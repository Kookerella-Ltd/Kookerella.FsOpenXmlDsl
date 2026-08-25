namespace Kookerella.FsOpenXmlDsl.Mcp

open System
open System.ComponentModel
open System.Globalization
open System.Text.Json
open System.Xml.Linq
open ModelContextProtocol.Server
open Kookerella.FsOpenXmlDsl

/// One sheet's worth of input for `WorkbookTools.CreateWorkbook`. A plain mutable class
/// (not an F# record) so `System.Text.Json`'s default reflection-based (de)serializer -
/// which the MCP SDK uses to turn a tool call's JSON arguments into .NET parameter values -
/// can construct one without needing an F#-aware converter: an F# record's single
/// all-fields constructor is exactly the shape `System.Text.Json` can't instantiate from
/// JSON without extra attributes, and using a plain class sidesteps that entirely rather
/// than depending on a workaround for it.
type SheetInput() =
    member val Name: string = "" with get, set
    member val Rows: string[][] = [||] with get, set

/// One cell of `WorkbookTools.ReadWorkbook`'s output - see `SheetInput`'s own doc comment
/// for why this is a plain class rather than a record.
type CellOutput() =
    member val Ref: string = "" with get, set
    member val Value: string = "" with get, set
    member val CachedValue: Nullable<float> = Nullable() with get, set

type SheetOutput() =
    member val Name: string = "" with get, set
    member val Cells: CellOutput[] = [||] with get, set

/// The MCP tool surface over `Kookerella.FsOpenXmlDsl`. `create_workbook`/`read_workbook`
/// are deliberately narrow - plain cell values and formulas only, addressed by a simple
/// row/column grid per sheet, no styling/tables/charts/pivot tables/etc. - the same
/// "honest, bounded MVP, documented gap" scoping this whole library uses elsewhere, not an
/// oversight. The other four tools (`generate_fsharp_script`/`generate_csharp_script`/
/// `generate_xml`/`create_workbook_from_xml`) aren't limited that way - they cover the full
/// worksheet/workbook-level feature set, just via generated source/XML rather than a plain
/// grid.
[<McpServerToolType>]
type WorkbookTools =

    /// Parses one grid cell's raw text into a `CellValue`, using the same conventions a
    /// human typing into Excel would expect: a leading "=" makes it a formula (with no
    /// cached value - see `Formula`'s own doc comment on why that's safe only when
    /// something downstream, e.g. Excel itself on open, will actually compute a result),
    /// "true"/"false" (case-insensitive) make a boolean, anything else that parses as a
    /// number is numeric, and everything else is plain text. Chosen over a more explicit
    /// tagged input shape (e.g. `{ Kind: "formula"; Text: "..." }`) because it's the
    /// natural way any LLM already renders spreadsheet content when asked, without needing
    /// this tool's schema explained in detail first.
    static member private ParseCellValue(raw: string) : CellValue =
        if String.IsNullOrEmpty raw then
            Empty
        elif raw.StartsWith("=") then
            Formula(raw.Substring(1), None)
        elif raw.Equals("true", StringComparison.OrdinalIgnoreCase) then
            Boolean true
        elif raw.Equals("false", StringComparison.OrdinalIgnoreCase) then
            Boolean false
        else
            match Double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, n -> Number n
            | false, _ -> Text raw

    /// The inverse of `ParseCellValue` - renders a cell's value back to the same textual
    /// convention (a leading "=" for formulas), plus the formula's cached value (if any)
    /// separately, since that doesn't fit naturally into the text form.
    static member private RenderCellValue(v: CellValue) : string * float option =
        match v with
        | Empty -> "", None
        | Text s -> s, None
        | Number n -> n.ToString(CultureInfo.InvariantCulture), None
        | Boolean b -> (if b then "true" else "false"), None
        | Date d -> d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), None
        | Formula(expr, cached) -> "=" + expr, cached

    static member private ToWorksheet(input: SheetInput) : Worksheet =
        let cells =
            input.Rows
            |> Array.mapi (fun rowIndex row ->
                row
                |> Array.mapi (fun colIndex raw ->
                    { Ref = CellRef.create rowIndex colIndex
                      Value = WorkbookTools.ParseCellValue raw
                      Style = None }))
            |> Array.collect id
            |> Array.filter (fun c -> c.Value <> Empty)
            |> Array.toList

        sheetOfCells input.Name cells

    [<McpServerTool(Name = "create_workbook")>]
    [<Description(
        "Creates a new Excel workbook (.xlsx) from a simple grid of sheets/rows/cells and saves it to disk. \
         Each cell is given as plain text, the same way you'd type it into Excel: a leading '=' makes it a \
         formula (e.g. \"=SUM(A1:A2)\"), 'true'/'false' makes a boolean, a bare number is numeric, and \
         anything else is text. Rows don't need to be the same length. Does not support cell styling, \
         tables, charts, or pivot tables in this version - reference the Kookerella.FsOpenXmlDsl library \
         directly for those."
    )>]
    static member CreateWorkbook
        (
            [<Description("Output file path, e.g. \"C:\\reports\\invoice.xlsx\". The directory must already exist.")>] path: string,
            [<Description("The sheets to create, in order. Each sheet has a Name and a Rows grid (array of rows, each an array of cell text).")>] sheets: SheetInput[]
        ) : string =
        let worksheets = sheets |> Array.map WorkbookTools.ToWorksheet |> Array.toList
        let wb = workbook worksheets
        Workbook.save path wb
        sprintf "Wrote %s (%d sheet%s)." path worksheets.Length (if worksheets.Length = 1 then "" else "s")

    [<McpServerTool(Name = "read_workbook")>]
    [<Description(
        "Reads an existing Excel workbook (.xlsx/.xlsm) and returns its sheets and cell contents as JSON. \
         Formula cells are rendered as \"=expression\" (matching create_workbook's input convention), with \
         any cached value included separately under CachedValue. Features outside the core cell model \
         (charts, tables, pivot tables, styling, etc.) are not included in this output - see MAPPING.md in \
         the main library repo for the full list of what round-trips."
    )>]
    static member ReadWorkbook([<Description("Path to an existing .xlsx or .xlsm file.")>] path: string) : string =
        let wb = Workbook.load path

        let sheets =
            wb.Sheets
            |> List.map (fun sheet ->
                let output = SheetOutput()
                output.Name <- sheet.Name

                output.Cells <-
                    sheet.Cells
                    |> List.map (fun cell ->
                        let text, cached = WorkbookTools.RenderCellValue cell.Value
                        let cellOutput = CellOutput()
                        cellOutput.Ref <- CellRef.toA1 cell.Ref
                        cellOutput.Value <- text
                        cellOutput.CachedValue <- (cached |> Option.map Nullable.op_Implicit |> Option.defaultValue (Nullable()))
                        cellOutput)
                    |> List.toArray

                output)
            |> List.toArray

        JsonSerializer.Serialize(sheets, JsonSerializerOptions(WriteIndented = true))

    [<McpServerTool(Name = "generate_fsharp_script")>]
    [<Description(
        "Reads an existing Excel workbook and returns a self-contained F# script (using Kookerella.FsOpenXmlDsl) \
         that rebuilds an equivalent file when run via `dotnet fsi`. Useful for explaining how a file is \
         structured, or as a starting point for a caller who wants the library's full feature set (styling, \
         tables, charts, pivot tables, etc.) beyond what create_workbook exposes."
    )>]
    static member GenerateFSharpScript
        (
            [<Description("Path to an existing .xlsx/.xlsm file to reverse-engineer into F# source.")>] path: string,
            [<Description("The output filename the generated script should save its rebuilt file to, e.g. \"output.xlsx\".")>] outputFileName: string
        ) : string =
        let wb = Workbook.load path

        let hashR (assemblyLocation: string) = sprintf "#r \"%s\"" (assemblyLocation.Replace("\\", "\\\\"))

        let referenceLines =
            [ hashR typeof<Workbook>.Assembly.Location
              hashR typeof<DocumentFormat.OpenXml.Packaging.SpreadsheetDocument>.Assembly.Location ]

        Workbook.generateScript referenceLines outputFileName wb

    [<McpServerTool(Name = "generate_csharp_script")>]
    [<Description(
        "Reads an existing Excel workbook and returns a self-contained C# file (using Kookerella.CsOpenXmlDsl) \
         that rebuilds an equivalent file when run via `dotnet run <file>.cs` (.NET 10's file-based apps \
         feature - no .csproj needed). The C# equivalent of generate_fsharp_script, for a caller who wants \
         pasteable/runnable C# rather than F# - useful for explaining how a file is structured, or as a \
         starting point for the wrapper's fluent API (styling, tables, charts, pivot tables, sparklines, \
         conditional formatting, data validation, hyperlinks, comments, print settings, defined names, \
         protection, etc. - Kookerella.CsOpenXmlDsl now covers the same worksheet/workbook-level feature \
         set as generate_fsharp_script's own Kookerella.FsOpenXmlDsl) beyond what create_workbook exposes."
    )>]
    static member GenerateCSharpScript
        (
            [<Description("Path to an existing .xlsx/.xlsm file to reverse-engineer into C# source.")>] path: string,
            [<Description("The output filename the generated script should save its rebuilt file to, e.g. \"output.xlsx\".")>] outputFileName: string
        ) : string =
        let wb = Kookerella.CsOpenXmlDsl.WorkbookIO.Load(path)

        // Unlike generate_fsharp_script's `#r` (a raw, machine-specific DLL path - .NET file-based
        // apps don't support `#r` at all, only `#:package`/`#:project`), this points at the published
        // NuGet package matching whatever version of the wrapper this server itself was built against -
        // portable to any machine with the .NET 10 SDK, not just this one.
        let packageVersion: Version = typeof<Kookerella.CsOpenXmlDsl.Workbook>.Assembly.GetName().Version

        let referenceLines =
            [| sprintf "#:package Kookerella.CsOpenXmlDsl@%d.%d.%d" packageVersion.Major packageVersion.Minor packageVersion.Build |]

        Kookerella.CsOpenXmlDsl.CsCodeGen.Generate(referenceLines, outputFileName, wb)

    [<McpServerTool(Name = "generate_xml")>]
    [<Description(
        "Reads an existing Excel workbook and returns it as XML, validated against Kookerella.FsOpenXmlDsl's \
         own embedded schema (Xml.xsd). A plain-data alternative to generate_fsharp_script/generate_csharp_script \
         for a caller who wants to inspect, transform (e.g. via XSLT), or archive a workbook's structure without \
         any F#/C# source involved - unlike those two, this returns data, not a runnable script, so there is no \
         output-filename parameter to control what a rebuild saves as. Covers the same worksheet/workbook-level \
         feature set generate_fsharp_script does."
    )>]
    static member GenerateXml([<Description("Path to an existing .xlsx/.xlsm file to convert to XML.")>] path: string) : string =
        let wb = Workbook.load path
        (Xml.ofWorkbook wb).ToString()

    [<McpServerTool(Name = "create_workbook_from_xml")>]
    [<Description(
        "Builds a new Excel workbook from XML matching Kookerella.FsOpenXmlDsl's own embedded schema (Xml.xsd) \
         and saves it to disk - the inverse of generate_xml. The natural target for a caller that already \
         produces data as XML (e.g. an XSLT pipeline generating a report) and wants to reach Excel without \
         learning the OOXML schema or this library's own F#/C# API. Covers the same worksheet/workbook-level \
         feature set generate_xml does; unlike create_workbook, this isn't limited to plain cell values - \
         styling, tables, charts, and every other modeled feature can be expressed in the XML."
    )>]
    static member CreateWorkbookFromXml
        (
            [<Description("The workbook XML content - a <workbook> root element matching Xml.xsd.")>] xml: string,
            [<Description("Output file path, e.g. \"C:\\reports\\invoice.xlsx\". The directory must already exist.")>] path: string
        ) : string =
        let wb = XElement.Parse(xml) |> Xml.toWorkbook
        Workbook.save path wb
        sprintf "Wrote %s (%d sheet%s)." path wb.Sheets.Length (if wb.Sheets.Length = 1 then "" else "s")
