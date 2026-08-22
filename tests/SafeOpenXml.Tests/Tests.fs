module SafeOpenXml.Tests

open System
open System.Diagnostics
open System.IO
open Xunit
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Validation
open SafeOpenXml
open type SafeOpenXml.SheetDsl

[<Theory>]
[<InlineData(0, 0, "A1")>]
[<InlineData(0, 25, "Z1")>]
[<InlineData(0, 26, "AA1")>]
[<InlineData(4, 27, "AB5")>]
[<InlineData(99, 701, "ZZ100")>]
let ``CellRef A1 round trip`` (row: int) (col: int) (expected: string) =
    let r = CellRef.create row col
    Assert.Equal(expected, CellRef.toA1 r)
    Assert.Equal<CellRef>(r, CellRef.ofA1 expected)

// --- Scenario harness -------------------------------------------------------------
//
// Each scenario below is a self-contained demonstration of one feature. Running it
// writes the workbook it builds to Examples/<scenario name>/output.xlsx (checked into
// the repo) so you can open any single feature in Excel without re-running anything,
// while the test itself verifies the file is schema-valid and round-trips exactly back
// through the DSL - the same two checks every prior test in this file already made,
// just factored out so each scenario only has to state what it's building.

let private examplesDir = Path.Combine(__SOURCE_DIRECTORY__, "Examples")

let private assertSchemaValid (path: string) =
    use document = SpreadsheetDocument.Open(path, false)
    let validator = OpenXmlValidator()
    let errors = validator.Validate(document) |> List.ofSeq

    Assert.True(
        errors.IsEmpty,
        String.Join("\n", errors |> Seq.map (fun e -> sprintf "%s: %s" e.Path.XPath e.Description))
    )

let private assertWorksheetRoundTrips (original: Worksheet) (path: string) =
    let roundTripped = Workbook.load path
    let actual = roundTripped.Sheets |> List.find (fun s -> s.Name = original.Name)

    Assert.Equal<Cell list>(
        original.Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col),
        actual.Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col)
    )

    Assert.Equal<Map<int, ColumnProps>>(original.ColumnProps, actual.ColumnProps)
    Assert.Equal<Map<int, RowProps>>(original.RowProps, actual.RowProps)
    Assert.Equal<MergedRange list>(original.MergedRanges, actual.MergedRanges)
    Assert.Equal<FreezePane option>(original.FreezePane, actual.FreezePane)
    Assert.Equal<AutoFilterRange option>(original.AutoFilter, actual.AutoFilter)

    // Password never round-trips (the hash isn't reversible - see SheetProtection's own
    // doc comment), so compare with it normalized away on both sides rather than skipping
    // this check entirely.
    Assert.Equal<SheetProtection option>(
        original.Protection |> Option.map (fun p -> { p with Password = None }),
        actual.Protection |> Option.map (fun p -> { p with Password = None })
    )
    Assert.Equal<ConditionalFormatEntry list>(original.ConditionalFormats, actual.ConditionalFormats)
    Assert.Equal<DataValidationEntry list>(original.DataValidations, actual.DataValidations)
    Assert.Equal<HyperlinkEntry list>(original.Hyperlinks, actual.Hyperlinks)
    Assert.Equal<CommentEntry list>(original.Comments, actual.Comments)
    Assert.Equal<PageSetup option>(original.PageSetup, actual.PageSetup)
    Assert.Equal<TableEntry list>(original.Tables, actual.Tables)

/// An F# string literal can't contain a raw backslash, so a Windows assembly path needs
/// its separators doubled before it's safe to splice into a generated `#r "..."` line.
let private hashR (assemblyLocation: string) =
    sprintf "#r \"%s\"" (assemblyLocation.Replace("\\", "\\\\"))

/// `dotnet fsi` needs `#r` for both SafeOpenXml itself and its OpenXml SDK dependency -
/// resolved from whatever assemblies this very test run already loaded, so the generated
/// scripts work regardless of build configuration (Debug/Release) or where the repo lives.
let private codeGenReferenceLines =
    [ hashR typeof<Workbook>.Assembly.Location
      hashR typeof<SpreadsheetDocument>.Assembly.Location ]

/// Saves `wb` to `Examples/<name>/output.xlsx`, asserts the file is schema-valid, and
/// asserts every sheet round-trips exactly back through the DSL. Also writes an
/// `Examples/<name>/script.fsx` that regenerates the same file - see the `Category=Slow`
/// tests below for where that script actually gets executed and verified.
let private verifyScenario (name: string) (wb: Workbook) =
    let dir = Path.Combine(examplesDir, name)
    Directory.CreateDirectory(dir) |> ignore
    let path = Path.Combine(dir, "output.xlsx")
    Workbook.save path wb

    assertSchemaValid path
    wb.Sheets |> List.iter (fun sheet -> assertWorksheetRoundTrips sheet path)

    let roundTripped = Workbook.load path
    Assert.Equal<DefinedNameEntry list>(wb.DefinedNames, roundTripped.DefinedNames)

    let script = Workbook.generateScript codeGenReferenceLines "output.xlsx" wb
    File.WriteAllText(Path.Combine(dir, "script.fsx"), script)

// --- Core: cell values, styles, layout --------------------------------------------

let private headerStyle: CellStyle =
    { CellStyle.Default with
        Font = Some { FontStyle.Default with Bold = true; Color = Some Color.white }
        Fill = Some { Color = Rgb(68uy, 84uy, 106uy) }
        Border =
            Some
                { BorderStyle.None with
                    Bottom = Some { Style = Thin; Color = Some Color.black } } }

[<Fact>]
let ``example: text number boolean date cells`` () =
    // A Date cell needs an explicit number format if you want the round-tripped Style to
    // stay `None`: with no style at all, the writer auto-applies ShortDate itself (see
    // Writer.fs) so Excel shows a date instead of a raw OLE Automation serial number.
    let dateStyle = { CellStyle.Default with NumberFormat = Some ShortDate }

    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Name")
                    cell (Number 42.5)
                    cell (Boolean true)
                    cell (Date(DateTime(2026, 3, 1)), style = dateStyle) ] ]

    verifyScenario "TextNumberBooleanDateCells" (workbook [ data ])

[<Fact>]
let ``example: styled header row`` () =
    // See the comment in ``example: text number boolean date cells`` re: why the Date
    // cell needs an explicit style here.
    let dateStyle = { CellStyle.Default with NumberFormat = Some ShortDate }

    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Name", style = headerStyle)
                    cell (Text "Amount", style = headerStyle)
                    cell (Text "Purchased", style = headerStyle) ]
              row [ cell (Text "Widgets"); cell (Number 42.5); cell (Date(DateTime(2026, 1, 15)), style = dateStyle) ] ]

    verifyScenario "StyledHeaderRow" (workbook [ data ])

[<Fact>]
let ``example: number formats`` () =
    let styled nf value = cell (value, style = { CellStyle.Default with NumberFormat = Some nf })

    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Integer"); styled Integer (Number 1234.0) ]
              row [ cell (Text "TwoDecimal"); styled TwoDecimal (Number 1234.5) ]
              row [ cell (Text "Percentage"); styled Percentage (Number 0.42) ]
              row [ cell (Text "Currency"); styled Currency (Number 19.99) ]
              row [ cell (Text "ShortDate"); styled ShortDate (Date(DateTime(2026, 3, 1))) ]
              row [ cell (Text "DateAndTime"); styled DateAndTime (Date(DateTime(2026, 3, 1, 13, 30, 0))) ] ]

    verifyScenario "NumberFormats" (workbook [ data ])

[<Fact>]
let ``example: formulas`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Widgets"); cell (Number 4.0); cell (Number 2.5); cell (Formula("B1*C1", Some 10.0)) ]
              row [ cell (Text "Gadgets"); cell (Number 2.0); cell (Number 19.99); cell (Formula("B2*C2", Some 39.98)) ]
              row [ cell (Text "Total"); cell (Formula("SUM(D1:D2)", Some 49.98)) ] ]

    verifyScenario "Formulas" (workbook [ data ])

[<Fact>]
let ``example: merged cells`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Quarterly Report", style = headerStyle) ]
              Merge(CellRef.ofA1 "A1", CellRef.ofA1 "D1")
              row [ cell (Text "Q1"); cell (Text "Q2"); cell (Text "Q3"); cell (Text "Q4") ] ]

    verifyScenario "MergedCells" (workbook [ data ])

[<Fact>]
let ``example: frozen panes`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Name", style = headerStyle); cell (Text "Amount", style = headerStyle) ]
              row [ cell (Text "Row 1"); cell (Number 1.0) ]
              row [ cell (Text "Row 2"); cell (Number 2.0) ]
              Freeze(1, 0) ]

    verifyScenario "FrozenPanes" (workbook [ data ])

[<Fact>]
let ``example: column widths and row heights`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Wide column"); cell (Text "Tall row") ]
              ColumnWidth(0, 30.0)
              RowHeight(0, 30.0) ]

    verifyScenario "ColumnWidthsAndRowHeights" (workbook [ data ])

[<Fact>]
let ``example: sparse row and column jumps`` () =
    // `row`/`cell` advance sequentially by default; `index =`/`col =` jump to an explicit
    // position and sequential numbering resumes right after it - this leaves row 2 (index 1)
    // and column B empty, landing values at A1, A3, and C4.
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "A1") ]
              row (
                  [ cell (Text "A3", col = 0); cell (Text "C4", col = 2) ],
                  index = 2
              ) ]

    verifyScenario "SparseRowAndColumnJumps" (workbook [ data ])

[<Fact>]
let ``example: multiple sheets`` () =
    let wb =
        workbook
            [ sheetOfCells "First" [ cellA1 "A1" (Text "one") ]
              sheetOfCells "Second" [ cellA1 "A1" (Number 2.0) ] ]

    verifyScenario "MultipleSheets" wb
    Assert.Equal<string list>([ "First"; "Second" ], wb.Sheets |> List.map (fun s -> s.Name))

// --- Conditional formatting --------------------------------------------------------

let private redFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(255uy, 199uy, 206uy) } }
let private greenFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(198uy, 239uy, 206uy) } }

[<Fact>]
let ``example: conditional format cell value rule`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 50.0) ]
              row [ cell (Number 150.0) ]
              row [ cell (Number 90.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", CellValueRule(GreaterThan, "100", None, redFillStyle)) ]

    verifyScenario "ConditionalFormat_CellValueRule" (workbook [ data ])

[<Fact>]
let ``example: conditional format formula rule`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 10.0); cell (Number 20.0) ]
              row [ cell (Number 30.0); cell (Number 5.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A2", FormulaRule("A1>B1", greenFillStyle)) ]

    verifyScenario "ConditionalFormat_FormulaRule" (workbook [ data ])

[<Fact>]
let ``example: conditional format 2-color scale`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 10.0) ]
              row [ cell (Number 50.0) ]
              row [ cell (Number 90.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", ColorScale2(Color.white, Color.red)) ]

    verifyScenario "ConditionalFormat_ColorScale2" (workbook [ data ])

[<Fact>]
let ``example: conditional format 3-color scale`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 10.0) ]
              row [ cell (Number 50.0) ]
              row [ cell (Number 90.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", ColorScale3(Color.red, Color.yellow, Color.green)) ]

    verifyScenario "ConditionalFormat_ColorScale3" (workbook [ data ])

[<Fact>]
let ``example: conditional format data bar`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 10.0) ]
              row [ cell (Number 50.0) ]
              row [ cell (Number 90.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", DataBarRule Color.blue) ]

    verifyScenario "ConditionalFormat_DataBar" (workbook [ data ])

[<Fact>]
let ``example: conditional format duplicate values`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Apple") ]
              row [ cell (Text "Banana") ]
              row [ cell (Text "Apple") ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", DuplicateValuesRule redFillStyle) ]

    verifyScenario "ConditionalFormat_DuplicateValues" (workbook [ data ])

[<Fact>]
let ``example: conditional format unique values`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Apple") ]
              row [ cell (Text "Banana") ]
              row [ cell (Text "Apple") ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A3", UniqueValuesRule greenFillStyle) ]

    verifyScenario "ConditionalFormat_UniqueValues" (workbook [ data ])

// --- Data validation ----------------------------------------------------------------

[<Fact>]
let ``example: data validation inline list`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Size") ]
              dataValidation (CellRef.ofA1 "A2", CellRef.ofA1 "A2", ListValidation [ "Small"; "Medium"; "Large" ]) ]

    verifyScenario "DataValidation_List" (workbook [ data ])

[<Fact>]
let ``example: data validation list from range`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Small"); cell (Text "Medium"); cell (Text "Large") ]
              row [ cell (Text "Size") ]
              dataValidation (CellRef.ofA1 "A2", CellRef.ofA1 "A2", ListFromRangeValidation(CellRef.ofA1 "A1", CellRef.ofA1 "C1")) ]

    verifyScenario "DataValidation_ListFromRange" (workbook [ data ])

[<Fact>]
let ``example: data validation whole number`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Quantity") ]
              dataValidation (
                  CellRef.ofA1 "A2",
                  CellRef.ofA1 "A2",
                  WholeNumberValidation(GreaterThan, "0", None),
                  errorTitle = "Invalid quantity",
                  errorMessage = "Quantity must be a positive whole number."
              ) ]

    verifyScenario "DataValidation_WholeNumber" (workbook [ data ])

[<Fact>]
let ``example: data validation decimal`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Fraction (0-1)") ]
              dataValidation (CellRef.ofA1 "A2", CellRef.ofA1 "A2", DecimalValidation(Between, "0", Some "1")) ]

    verifyScenario "DataValidation_Decimal" (workbook [ data ])

[<Fact>]
let ``example: data validation text length`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Short code (<= 10 chars)") ]
              dataValidation (CellRef.ofA1 "A2", CellRef.ofA1 "A2", TextLengthValidation(LessThanOrEqual, "10", None)) ]

    verifyScenario "DataValidation_TextLength" (workbook [ data ])

[<Fact>]
let ``example: data validation custom formula`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Must be a number") ]
              dataValidation (
                  CellRef.ofA1 "A2",
                  CellRef.ofA1 "A2",
                  CustomValidation("ISNUMBER(A2)"),
                  allowBlank = false,
                  inputTitle = "Note",
                  inputMessage = "Enter a numeric value."
              ) ]

    verifyScenario "DataValidation_Custom" (workbook [ data ])

// --- Hyperlinks ---------------------------------------------------------------------

[<Fact>]
let ``example: hyperlink external`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Open-XML-SDK on GitHub") ]
              hyperlink (
                  CellRef.ofA1 "A1",
                  ExternalHyperlink "https://github.com/dotnet/Open-XML-SDK",
                  tooltip = "Open in browser",
                  display = "dotnet/Open-XML-SDK"
              ) ]

    verifyScenario "Hyperlink_External" (workbook [ data ])

[<Fact>]
let ``example: hyperlink internal`` () =
    let sheet1 =
        sheet
            "Sheet1"
            [ row [ cell (Text "Go to Sheet2") ]
              hyperlink (CellRef.ofA1 "A1", InternalHyperlink "Sheet2!A1") ]

    let sheet2 = sheet "Sheet2" [ row [ cell (Text "You made it!") ] ]

    verifyScenario "Hyperlink_Internal" (workbook [ sheet1; sheet2 ])

// --- Comments -------------------------------------------------------------------------

[<Fact>]
let ``example: comments`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Revenue"); cell (Number 1250.0) ]
              row [ cell (Text "Costs"); cell (Number 900.0) ]
              comment (CellRef.ofA1 "B1", "Figure is provisional pending audit.", author = "Alex")
              comment (CellRef.ofA1 "B2", "Includes one-off relocation costs.", author = "Alex")
              comment (CellRef.ofA1 "A1", "Double check this label.") ]

    verifyScenario "Comments" (workbook [ data ])

[<Fact>]
let ``comments VML drawing part is well-formed XML`` () =
    // vmlDrawingContent is hand-templated raw XML (see the comment in Writer.fs on why -
    // VML predates OOXML's typed object model), so unlike everything else this feature
    // touches, it isn't schema-checked by assertSchemaValid. This is the narrower check
    // that actually applies to it: independent of any other test, confirm the string it
    // produces at least parses as well-formed XML.
    let data = sheet "Sheet1" [ row [ cell (Text "X") ]; comment (CellRef.ofA1 "A1", "A note") ]
    use stream = new MemoryStream()
    Workbook.saveToStream stream (workbook [ data ])
    stream.Position <- 0L

    use document = SpreadsheetDocument.Open(stream, false)
    let worksheetPart = document.WorkbookPart.WorksheetParts |> Seq.head
    let vmlPart = worksheetPart.VmlDrawingParts |> Seq.head
    use reader = new StreamReader(vmlPart.GetStream())
    let content = reader.ReadToEnd()
    let xml = System.Xml.Linq.XDocument.Parse(content)
    Assert.NotNull(xml.Root)

// --- AutoFilter -----------------------------------------------------------------------

[<Fact>]
let ``example: autofilter`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Name", style = headerStyle)
                    cell (Text "Amount", style = headerStyle)
                    cell (Text "Region", style = headerStyle) ]
              row [ cell (Text "Widgets"); cell (Number 42.5); cell (Text "North") ]
              row [ cell (Text "Gadgets"); cell (Number 19.99); cell (Text "South") ]
              autoFilter (CellRef.ofA1 "A1", CellRef.ofA1 "C3") ]

    verifyScenario "AutoFilter" (workbook [ data ])

// --- Protection -----------------------------------------------------------------------

[<Fact>]
let ``example: cell locking`` () =
    // Unlocked so it stays editable once the sheet is protected; the label cell keeps the
    // implicit Locked = true default (Excel locks every cell unless told otherwise).
    let unlocked = { CellStyle.Default with Protection = Some { Locked = false; Hidden = false } }

    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Enter quantity:"); cell (Number 0.0, style = unlocked) ]
              Protect SheetProtection.Default ]

    verifyScenario "CellLocking" (workbook [ data ])

[<Fact>]
let ``example: sheet protection with password`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Protected sheet") ]
              Protect
                  { SheetProtection.Default with
                      Password = Some "hunter2"
                      FormatCells = Some true
                      Sort = Some true
                      AutoFilter = Some true } ]

    verifyScenario "SheetProtectionWithPassword" (workbook [ data ])

// --- Defined names ----------------------------------------------------------------------

[<Fact>]
let ``example: defined names`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number 0.075) ]
              row [ cell (Number 100.0) ]
              row [ cell (Formula("B1*(1+TaxRate)", Some 107.5)) ] ]

    let wb =
        workbook [ data ]
        |> withDefinedNames
            [ definedName "TaxRate" "Sheet1!$A$1"
              sheetScopedDefinedName "Sheet1" "LocalTotal" "Sheet1!$A$2" ]

    verifyScenario "DefinedNames" wb

// --- Print settings and page setup -------------------------------------------------------

[<Fact>]
let ``example: page setup landscape with margins and header footer`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Wide report") ]
              PageSetup
                  { PageSetup.Default with
                      Orientation = Landscape
                      PaperSize = Some A4
                      Scaling = Some(ScalePercent 85)
                      Margins = { PageMargins.Default with Left = 0.5; Right = 0.5 }
                      Header = Some "&C&\"Arial,Bold\"Quarterly Report"
                      Footer = Some "&LPage &P of &N&R&D" } ]

    verifyScenario "PageSetupLandscapeWithMargins" (workbook [ data ])

[<Fact>]
let ``example: page setup fit to one page wide`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Fits to one page wide") ]
              PageSetup { PageSetup.Default with Scaling = Some(FitToPage(1, 0)) } ]

    verifyScenario "PageSetupFitToOnePageWide" (workbook [ data ])

// --- Tables --------------------------------------------------------------------------

[<Fact>]
let ``example: table`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Item"); cell (Text "Quantity") ]
              row [ cell (Text "Widgets"); cell (Number 12.0) ]
              row [ cell (Text "Gadgets"); cell (Number 5.0) ]
              Table
                  { TopLeft = CellRef.ofA1 "A1"
                    BottomRight = CellRef.ofA1 "B3"
                    Name = "Inventory"
                    Columns = [ { Name = "Item"; CalculatedFormula = None }; { Name = "Quantity"; CalculatedFormula = None } ]
                    Style = TableStyle.Default } ]

    verifyScenario "Table" (workbook [ data ])

[<Fact>]
let ``example: table with calculated column and custom style`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Quantity"); cell (Text "Unit Price"); cell (Text "Total") ]
              row [ cell (Number 12.0); cell (Number 2.5); cell (Formula("[@Quantity]*[@[Unit Price]]", Some 30.0)) ]
              row [ cell (Number 5.0); cell (Number 9.0); cell (Formula("[@Quantity]*[@[Unit Price]]", Some 45.0)) ]
              Table
                  { TopLeft = CellRef.ofA1 "A1"
                    BottomRight = CellRef.ofA1 "C3"
                    Name = "Orders"
                    Columns =
                      [ { Name = "Quantity"; CalculatedFormula = None }
                        { Name = "Unit Price"; CalculatedFormula = None }
                        { Name = "Total"
                          CalculatedFormula = Some "[@Quantity]*[@[Unit Price]]" } ]
                    Style =
                      { TableStyle.Default with
                          Name = Some "TableStyleLight9"
                          ShowColumnStripes = true
                          ShowRowStripes = false } } ]

    verifyScenario "TableWithCalculatedColumn" (workbook [ data ])

// --- Generated-script verification (slow: actually runs `dotnet fsi`) -------------------
//
// Every scenario above writes its own Examples/<name>/script.fsx as a side effect of
// `verifyScenario`. These tests are the only place that script actually gets *executed*
// rather than just generated - each one runs its scenario's script via `dotnet fsi` and
// checks the regenerated output.xlsx round-trips to the same Workbook value as the
// committed one. Running `dotnet fsi` from cold is slow (multi-second startup per
// process), so this is its own Category=Slow group rather than part of the default
// `dotnet test` loop - run it explicitly with:
//   dotnet test --filter "Category=Slow"
// The default fast loop is:
//   dotnet test --filter "Category!=Slow"

/// One `obj[]` per scenario folder that has a `script.fsx` (i.e. every scenario above, once
/// its fast test has run at least once) - discovered from disk rather than hand-listed, so
/// a new scenario is automatically covered without touching this file.
let scenarioNames : obj[] seq =
    if Directory.Exists examplesDir then
        Directory.GetDirectories examplesDir
        |> Seq.filter (fun dir -> File.Exists(Path.Combine(dir, "script.fsx")))
        |> Seq.map (fun dir -> [| box (Path.GetFileName dir) |])
    else
        Seq.empty

[<Theory>]
[<Trait("Category", "Slow")>]
[<MemberData(nameof scenarioNames)>]
let ``example script regenerates an equivalent file`` (name: string) =
    let dir = Path.Combine(examplesDir, name)
    let scriptPath = Path.Combine(dir, "script.fsx")
    let outputPath = Path.Combine(dir, "output.xlsx")

    let before = Workbook.load outputPath

    // The OpenXml SDK's read side can leave a memory-mapped view of the file lingering
    // past `Dispose` on Windows (a documented SDK quirk, not something under this
    // library's control) - without forcing it closed here, the `dotnet fsi` subprocess
    // below can fail to reopen the same path with "a file with a user-mapped section open".
    GC.Collect()
    GC.WaitForPendingFinalizers()

    use proc =
        Process.Start(
            ProcessStartInfo(
                FileName = "dotnet",
                Arguments = sprintf "fsi \"%s\"" scriptPath,
                WorkingDirectory = dir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            )
        )

    let stdout = proc.StandardOutput.ReadToEnd()
    let stderr = proc.StandardError.ReadToEnd()
    proc.WaitForExit()

    Assert.True(proc.ExitCode = 0, sprintf "dotnet fsi %s failed (exit %d):\n%s\n%s" name proc.ExitCode stdout stderr)

    let after = Workbook.load outputPath
    Assert.Equal<Workbook>(before, after)
