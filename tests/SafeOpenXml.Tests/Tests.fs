module SafeOpenXml.Tests

open System
open System.Diagnostics
open System.IO
open Xunit
open DocumentFormat.OpenXml
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

    if original.PivotTables.IsEmpty then
        Assert.Equal<Cell list>(
            original.Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col),
            actual.Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col)
        )
    else
        // A pivot table injects its computed grid into the sheet's cells at write time -
        // see `PivotTableEntry`'s own doc comment - so `actual.Cells` legitimately has
        // more entries than `original` ever specified. Assert every originally-authored
        // cell survived unchanged instead of exact equality; the pivot table tests
        // themselves separately assert the computed grid's own values are correct.
        let actualByRef = actual.Cells |> List.map (fun c -> c.Ref, c) |> Map.ofList

        for cell in original.Cells do
            match actualByRef.TryFind cell.Ref with
            | Some actualCell -> Assert.Equal<Cell>(cell, actualCell)
            | None -> Assert.Fail(sprintf "Expected authored cell at %s to survive, but it's missing" (CellRef.toA1 cell.Ref))

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
    Assert.Equal<SparklineGroupEntry list>(original.SparklineGroups, actual.SparklineGroups)
    Assert.Equal<ChartEntry list>(original.Charts, actual.Charts)
    // F#'s structural equality on records compares `byte[]` fields by content, not
    // reference, so this is a genuine byte-for-byte comparison of the embedded image data.
    Assert.Equal<ImageEntry list>(original.Images, actual.Images)
    Assert.Equal<PivotTableEntry list>(original.PivotTables, actual.PivotTables)

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

/// Saves `wb` to `Examples/<name>/<fileName>`, asserts the file is schema-valid, and
/// asserts every sheet round-trips exactly back through the DSL. Also writes an
/// `Examples/<name>/script.fsx` that regenerates the same file - see the `Category=Slow`
/// tests below for where that script actually gets executed and verified. `fileName`
/// defaults to `output.xlsx`; the one exception is a workbook carrying a VBA project,
/// which needs the `.xlsm` extension real Excel expects for macro-enabled content.
let private verifyScenarioNamed (name: string) (fileName: string) (wb: Workbook) =
    let dir = Path.Combine(examplesDir, name)
    Directory.CreateDirectory(dir) |> ignore
    let path = Path.Combine(dir, fileName)
    Workbook.save path wb

    assertSchemaValid path
    wb.Sheets |> List.iter (fun sheet -> assertWorksheetRoundTrips sheet path)

    let roundTripped = Workbook.load path
    Assert.Equal<DefinedNameEntry list>(wb.DefinedNames, roundTripped.DefinedNames)

    // Password never round-trips (see WorkbookProtection's own doc comment), same
    // normalize-rather-than-skip approach as `assertWorksheetRoundTrips`'s sheet-level
    // protection check.
    Assert.Equal<WorkbookProtection option>(
        wb.Protection |> Option.map (fun p -> { p with Password = None }),
        roundTripped.Protection |> Option.map (fun p -> { p with Password = None })
    )

    // The VBA project is opaque bytes (see `Workbook.VbaProject`'s own doc comment) - F#'s
    // structural equality on `option`/array values compares by content, not reference, the
    // same as `ImageEntry.Data` above.
    Assert.Equal<byte[] option>(wb.VbaProject, roundTripped.VbaProject)

    let script = Workbook.generateScript codeGenReferenceLines fileName wb
    File.WriteAllText(Path.Combine(dir, "script.fsx"), script)

let private verifyScenario (name: string) (wb: Workbook) = verifyScenarioNamed name "output.xlsx" wb

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

/// A cell within one row, addressed by an A1-style reference rather than by position -
/// used only by the shared-formula test below since it hand-builds rows out of order
/// (formula cells in column A, value cells in column B) and needs each `Spreadsheet.Cell`
/// to carry an explicit `CellReference`, unlike the DSL's own `cell`/`row`.
let private rawCell (a1: string) (formula: Spreadsheet.CellFormula option) (cachedText: string) : Spreadsheet.Cell =
    let c = Spreadsheet.Cell(CellReference = StringValue(a1))
    formula |> Option.iter (fun f -> c.CellFormula <- f)
    c.CellValue <- Spreadsheet.CellValue(cachedText)
    c

let private sharedFormulaCell (sharedIndex: uint32) (masterText: string option) (reference: string option) : Spreadsheet.CellFormula =
    let f = Spreadsheet.CellFormula()
    f.FormulaType <- EnumValue<Spreadsheet.CellFormulaValues>(Spreadsheet.CellFormulaValues.Shared)
    f.SharedIndex <- UInt32Value(sharedIndex)
    masterText |> Option.iter (fun t -> f.Text <- t)
    reference |> Option.iter (fun r -> f.Reference <- StringValue(r))
    f

/// SafeOpenXml's own `Writer` never emits Excel's "shared formula" optimization (it always
/// writes each cell's own normal formula text verbatim) - there's no DSL construct that
/// produces one, so this can't be a `verifyScenario` gallery entry like every test above.
/// It's a real, common shape in files Excel itself saves (its default when you fill/drag a
/// formula across a range: only the group's first cell carries the actual expression text,
/// every other cell in the group carries just a cached value and a shared-group index), so
/// this hand-builds one directly against the OpenXml SDK - bypassing the DSL entirely, the
/// same way a genuinely foreign file would look - and checks `Workbook.load` reconstructs
/// each cell's own correctly-shifted formula text. See `Reader.fs`'s `formulaRefPattern`
/// and `shiftFormula` for the translation logic this exercises.
[<Fact>]
let ``reading a foreign file's shared formulas reconstructs each cell's own formula`` () =
    let path = Path.Combine(Path.GetTempPath(), sprintf "SafeOpenXmlSharedFormulaTest_%s.xlsx" (Guid.NewGuid().ToString("N")))

    try
        do
            use document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
            let workbookPart = document.AddWorkbookPart()
            workbookPart.Workbook <- Spreadsheet.Workbook()
            let sheets = Spreadsheet.Sheets()
            workbookPart.Workbook.AppendChild(sheets) |> ignore

            let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
            let sheetData = Spreadsheet.SheetData()
            worksheetPart.Worksheet <- Spreadsheet.Worksheet()
            worksheetPart.Worksheet.AppendChild(sheetData) |> ignore

            sheets.AppendChild(
                Spreadsheet.Sheet(
                    Name = StringValue("Sheet1"),
                    SheetId = UInt32Value(1u),
                    Id = StringValue(workbookPart.GetIdOfPart(worksheetPart))
                )
            )
            |> ignore

            let buildRow (rowIndex: int) (cells: Spreadsheet.Cell list) =
                let row = Spreadsheet.Row(RowIndex = UInt32Value(uint32 rowIndex))
                cells |> List.iter (fun c -> row.AppendChild(c) |> ignore)
                row

            // Group 0 (unanchored): A1 = SUM(B1:B1), filled down to A2/A3.
            let group0Master = sharedFormulaCell 0u (Some "SUM(B1:B1)") (Some "A1:A3")
            let group0Slave2 = sharedFormulaCell 0u None None
            let group0Slave3 = sharedFormulaCell 0u None None

            // Group 1 (mixed anchoring): A4 = SUM($B$1:B1), filled down to A5 -
            // the absolute endpoint must stay put while the relative one shifts.
            let group1Master = sharedFormulaCell 1u (Some "SUM($B$1:B1)") (Some "A4:A5")
            let group1Slave5 = sharedFormulaCell 1u None None

            // Group 2 (string literal): A6 = IF(B1=1,"A1","no"), filled down to A7 - the
            // quoted "A1" must survive untouched even though it looks exactly like a
            // reference, while the real reference B1 still shifts to B2.
            let group2Master = sharedFormulaCell 2u (Some "IF(B1=1,\"A1\",\"no\")") (Some "A6:A7")
            let group2Slave7 = sharedFormulaCell 2u None None

            sheetData.AppendChild(buildRow 1 [ rawCell "A1" (Some group0Master) "1"; rawCell "B1" None "10" ])
            |> ignore
            sheetData.AppendChild(buildRow 2 [ rawCell "A2" (Some group0Slave2) "2"; rawCell "B2" None "20" ])
            |> ignore
            sheetData.AppendChild(buildRow 3 [ rawCell "A3" (Some group0Slave3) "3"; rawCell "B3" None "30" ])
            |> ignore
            sheetData.AppendChild(buildRow 4 [ rawCell "A4" (Some group1Master) "4" ]) |> ignore
            sheetData.AppendChild(buildRow 5 [ rawCell "A5" (Some group1Slave5) "5" ]) |> ignore
            sheetData.AppendChild(buildRow 6 [ rawCell "A6" (Some group2Master) "0" ]) |> ignore
            sheetData.AppendChild(buildRow 7 [ rawCell "A7" (Some group2Slave7) "0" ]) |> ignore

            worksheetPart.Worksheet.Save()
            workbookPart.Workbook.Save()

        let wb = Workbook.load path
        let sheet1 = wb.Sheets |> List.find (fun s -> s.Name = "Sheet1")
        let cellAt (a1: string) = sheet1.Cells |> List.find (fun c -> c.Ref = CellRef.ofA1 a1)

        Assert.Equal(Formula("SUM(B1:B1)", Some 1.0), (cellAt "A1").Value)
        Assert.Equal(Formula("SUM(B2:B2)", Some 2.0), (cellAt "A2").Value)
        Assert.Equal(Formula("SUM(B3:B3)", Some 3.0), (cellAt "A3").Value)
        Assert.Equal(Formula("SUM($B$1:B1)", Some 4.0), (cellAt "A4").Value)
        Assert.Equal(Formula("SUM($B$1:B2)", Some 5.0), (cellAt "A5").Value)
        Assert.Equal(Formula("IF(B1=1,\"A1\",\"no\")", Some 0.0), (cellAt "A6").Value)
        Assert.Equal(Formula("IF(B2=1,\"A1\",\"no\")", Some 0.0), (cellAt "A7").Value)
    finally
        if File.Exists path then
            File.Delete path

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

[<Fact>]
let ``example: workbook structure protection`` () =
    let data = sheet "Sheet1" [ row [ cell (Text "Sheets can't be added, removed, or renamed") ] ]

    let wb =
        workbook [ data ]
        |> withProtection { WorkbookProtection.Default with Password = Some "hunter2"; LockStructure = Some true }

    verifyScenario "WorkbookProtection" wb

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

[<Fact>]
let ``example: page setup print area with multiple ranges`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Included") ]
              row [ cell (Text "Excluded") ]
              row [ cell (Text "Also included") ]
              PageSetup
                  { PageSetup.Default with
                      PrintArea = [ (CellRef.ofA1 "A1", CellRef.ofA1 "A1"); (CellRef.ofA1 "A3", CellRef.ofA1 "A3") ] } ]

    verifyScenario "PageSetupPrintArea" (workbook [ data ])

[<Fact>]
let ``example: page setup first page and even page header footer`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Report") ]
              PageSetup
                  { PageSetup.Default with
                      Header = Some "&CPage &P"
                      FirstHeader = Some "&CCover Page"
                      EvenFooter = Some "&L&F" } ]

    verifyScenario "PageSetupHeaderFooterVariants" (workbook [ data ])

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

// --- Sparklines ------------------------------------------------------------------------

[<Fact>]
let ``example: sparklines line group filled down`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Widgets"); cell (Number 3.0); cell (Number 8.0); cell (Number 5.0); cell (Number 9.0) ]
              row [ cell (Text "Gadgets"); cell (Number 6.0); cell (Number 4.0); cell (Number 7.0); cell (Number 2.0) ]
              SparklineGroup
                  { Style = { SparklineStyle.Default with ShowHigh = true; ShowLow = true }
                    Sparklines =
                      [ { Cell = CellRef.ofA1 "F1"; DataTopLeft = CellRef.ofA1 "B1"; DataBottomRight = CellRef.ofA1 "E1" }
                        { Cell = CellRef.ofA1 "F2"; DataTopLeft = CellRef.ofA1 "B2"; DataBottomRight = CellRef.ofA1 "E2" } ] } ]

    verifyScenario "SparklinesLineGroup" (workbook [ data ])

[<Fact>]
let ``example: sparklines column group with custom color`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Number -2.0); cell (Number 4.0); cell (Number -1.0); cell (Number 3.0) ]
              SparklineGroup
                  { Style =
                      { SparklineStyle.Default with
                          Type = Column
                          Color = Some(Rgb(0uy, 112uy, 192uy))
                          ShowNegative = true }
                    Sparklines = [ { Cell = CellRef.ofA1 "E1"; DataTopLeft = CellRef.ofA1 "A1"; DataBottomRight = CellRef.ofA1 "D1" } ] } ]

    verifyScenario "SparklinesColumnGroup" (workbook [ data ])

// --- Charts ------------------------------------------------------------------------------

[<Fact>]
let ``example: column chart with title and legend`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Quarter"); cell (Text "North"); cell (Text "South") ]
              row [ cell (Text "Q1"); cell (Number 12.0); cell (Number 9.0) ]
              row [ cell (Text "Q2"); cell (Number 15.0); cell (Number 11.0) ]
              row [ cell (Text "Q3"); cell (Number 9.0); cell (Number 14.0) ]
              EmbeddedChart
                  { Type = ChartColumn
                    Title = Some "Sales by Quarter"
                    CategoriesTopLeft = CellRef.ofA1 "A2"
                    CategoriesBottomRight = CellRef.ofA1 "A4"
                    Series =
                      [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B4" }
                        { Name = CellRef.ofA1 "C1"; ValuesTopLeft = CellRef.ofA1 "C2"; ValuesBottomRight = CellRef.ofA1 "C4" } ]
                    ShowLegend = true
                    TopLeftAnchor = CellRef.ofA1 "E1"
                    BottomRightAnchor = CellRef.ofA1 "L15" } ]

    verifyScenario "ChartColumn" (workbook [ data ])

[<Fact>]
let ``example: bar chart horizontal`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Team"); cell (Text "Score") ]
              row [ cell (Text "Alpha"); cell (Number 42.0) ]
              row [ cell (Text "Beta"); cell (Number 37.0) ]
              EmbeddedChart
                  { Type = ChartBar
                    Title = None
                    CategoriesTopLeft = CellRef.ofA1 "A2"
                    CategoriesBottomRight = CellRef.ofA1 "A3"
                    Series = [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B3" } ]
                    ShowLegend = false
                    TopLeftAnchor = CellRef.ofA1 "D1"
                    BottomRightAnchor = CellRef.ofA1 "K12" } ]

    verifyScenario "ChartBar" (workbook [ data ])

[<Fact>]
let ``example: line chart with two series`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Month"); cell (Text "2025"); cell (Text "2026") ]
              row [ cell (Text "Jan"); cell (Number 100.0); cell (Number 120.0) ]
              row [ cell (Text "Feb"); cell (Number 110.0); cell (Number 115.0) ]
              row [ cell (Text "Mar"); cell (Number 105.0); cell (Number 130.0) ]
              EmbeddedChart
                  { Type = ChartLine
                    Title = Some "Monthly Trend"
                    CategoriesTopLeft = CellRef.ofA1 "A2"
                    CategoriesBottomRight = CellRef.ofA1 "A4"
                    Series =
                      [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B4" }
                        { Name = CellRef.ofA1 "C1"; ValuesTopLeft = CellRef.ofA1 "C2"; ValuesBottomRight = CellRef.ofA1 "C4" } ]
                    ShowLegend = true
                    TopLeftAnchor = CellRef.ofA1 "E1"
                    BottomRightAnchor = CellRef.ofA1 "L15" } ]

    verifyScenario "ChartLine" (workbook [ data ])

[<Fact>]
let ``example: pie chart`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Segment"); cell (Text "Share") ]
              row [ cell (Text "Retail"); cell (Number 55.0) ]
              row [ cell (Text "Wholesale"); cell (Number 30.0) ]
              row [ cell (Text "Online"); cell (Number 15.0) ]
              EmbeddedChart
                  { Type = ChartPie
                    Title = Some "Revenue Share"
                    CategoriesTopLeft = CellRef.ofA1 "A2"
                    CategoriesBottomRight = CellRef.ofA1 "A4"
                    Series = [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B4" } ]
                    ShowLegend = true
                    TopLeftAnchor = CellRef.ofA1 "D1"
                    BottomRightAnchor = CellRef.ofA1 "K14" } ]

    verifyScenario "ChartPie" (workbook [ data ])

// --- Images --------------------------------------------------------------------------

/// The canonical "1x1 transparent GIF" - the smallest possible valid image file, used
/// ubiquitously as a web tracking pixel, so its bytes are about as well-known and
/// trustworthy as test fixtures get.
let private onePixelGif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7")

[<Fact>]
let ``example: image anchored over a range`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Logo below:") ]
              EmbeddedImage
                  { Data = onePixelGif
                    Format = Gif
                    TopLeftAnchor = CellRef.ofA1 "A3"
                    BottomRightAnchor = CellRef.ofA1 "C10" } ]

    verifyScenario "Image" (workbook [ data ])

[<Fact>]
let ``example: chart and image sharing one worksheet's drawing canvas`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Month"); cell (Text "Units") ]
              row [ cell (Text "Jan"); cell (Number 10.0) ]
              row [ cell (Text "Feb"); cell (Number 14.0) ]
              EmbeddedChart
                  { Type = ChartColumn
                    Title = None
                    CategoriesTopLeft = CellRef.ofA1 "A2"
                    CategoriesBottomRight = CellRef.ofA1 "A3"
                    Series = [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B3" } ]
                    ShowLegend = false
                    TopLeftAnchor = CellRef.ofA1 "D1"
                    BottomRightAnchor = CellRef.ofA1 "H8" }
              EmbeddedImage
                  { Data = onePixelGif
                    Format = Gif
                    TopLeftAnchor = CellRef.ofA1 "D10"
                    BottomRightAnchor = CellRef.ofA1 "F15" } ]

    verifyScenario "ChartAndImage" (workbook [ data ])

// --- Macros / VBA --------------------------------------------------------------------

/// A real `vbaProject.bin` - extracted from a workbook actually authored and saved by
/// Excel (a single standard module, `Sub HelloWorld()` writing to A1), not hand-built.
/// This DSL treats a VBA project as an opaque OLE/CFBF binary (see `Workbook.VbaProject`'s
/// own doc comment) - unlike every other fixture in this file, there's no reasonable way
/// to hand-construct a valid one, so this is checked in as a binary asset instead of an
/// inline base64 literal like `onePixelGif` above (it's a couple of orders of magnitude
/// bigger).
let private sampleVbaProject =
    File.ReadAllBytes(Path.Combine(__SOURCE_DIRECTORY__, "Assets", "sample.vbaProject.bin"))

[<Fact>]
let ``example: workbook with a vba macro`` () =
    let data = sheet "Sheet1" [ row [ cell (Text "Run the HelloWorld macro to fill A1") ] ]

    let wb = workbook [ data ] |> withVbaProject sampleVbaProject

    verifyScenarioNamed "VbaMacro" "output.xlsm" wb

// --- Pivot tables ------------------------------------------------------------------------
//
// Unlike every scenario above, a pivot table's correctness isn't fully captured by
// `PivotTableEntry` round-tripping - the actual point is the *computed* grid, so each
// test here also loads the saved file back and asserts specific aggregated cell values
// by hand, rather than relying only on `verifyScenario`'s generic equality check.

let private numberAt (wb: Workbook) (sheetName: string) (a1: string) : float =
    let sheet = wb.Sheets |> List.find (fun s -> s.Name = sheetName)
    let cellRef = CellRef.ofA1 a1

    match sheet.Cells |> List.tryFind (fun c -> c.Ref = cellRef) with
    | Some { Value = Number n } -> n
    | other -> failwithf "Expected a number at %s on '%s', got %A" a1 sheetName other

[<Fact>]
let ``example: pivot table row field only`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Region"); cell (Text "Sales") ]
              row [ cell (Text "East"); cell (Number 10.0) ]
              row [ cell (Text "West"); cell (Number 20.0) ]
              row [ cell (Text "East"); cell (Number 5.0) ]
              row [ cell (Text "West"); cell (Number 15.0) ]
              EmbeddedPivotTable
                  { SourceSheet = None
                    SourceTopLeft = CellRef.ofA1 "A1"
                    SourceBottomRight = CellRef.ofA1 "B5"
                    RowField = "Region"
                    ColumnField = None
                    ValueField = "Sales"
                    Aggregation = PivotSum
                    ValueCaption = None
                    TopLeftAnchor = CellRef.ofA1 "D1" } ]

    let scenarioName = "PivotTableRowOnly"
    verifyScenario scenarioName (workbook [ data ])

    let roundTripped = Workbook.load (Path.Combine(examplesDir, scenarioName, "output.xlsx"))
    Assert.Equal(15.0, numberAt roundTripped "Sheet1" "E2")
    Assert.Equal(35.0, numberAt roundTripped "Sheet1" "E3")
    Assert.Equal(50.0, numberAt roundTripped "Sheet1" "E4")

[<Fact>]
let ``example: pivot table row and column fields`` () =
    let data =
        sheet
            "Sheet1"
            [ row [ cell (Text "Region"); cell (Text "Quarter"); cell (Text "Sales") ]
              row [ cell (Text "East"); cell (Text "Q1"); cell (Number 10.0) ]
              row [ cell (Text "East"); cell (Text "Q2"); cell (Number 5.0) ]
              row [ cell (Text "West"); cell (Text "Q1"); cell (Number 20.0) ]
              row [ cell (Text "West"); cell (Text "Q2"); cell (Number 15.0) ]
              EmbeddedPivotTable
                  { SourceSheet = None
                    SourceTopLeft = CellRef.ofA1 "A1"
                    SourceBottomRight = CellRef.ofA1 "C5"
                    RowField = "Region"
                    ColumnField = Some "Quarter"
                    ValueField = "Sales"
                    Aggregation = PivotSum
                    ValueCaption = Some "Total Sales"
                    TopLeftAnchor = CellRef.ofA1 "E1" } ]

    let scenarioName = "PivotTableRowAndColumn"
    verifyScenario scenarioName (workbook [ data ])

    let roundTripped = Workbook.load (Path.Combine(examplesDir, scenarioName, "output.xlsx"))
    // E1 Region | F1 Q1 | G1 Q2 | H1 Grand Total
    // E2 East   | F2 10 | G2 5  | H2 15
    // E3 West   | F3 20 | G3 15 | H3 35
    // E4 Grand Total | F4 30 | G4 20 | H4 50
    Assert.Equal(10.0, numberAt roundTripped "Sheet1" "F2")
    Assert.Equal(5.0, numberAt roundTripped "Sheet1" "G2")
    Assert.Equal(15.0, numberAt roundTripped "Sheet1" "H2")
    Assert.Equal(20.0, numberAt roundTripped "Sheet1" "F3")
    Assert.Equal(15.0, numberAt roundTripped "Sheet1" "G3")
    Assert.Equal(35.0, numberAt roundTripped "Sheet1" "H3")
    Assert.Equal(30.0, numberAt roundTripped "Sheet1" "F4")
    Assert.Equal(20.0, numberAt roundTripped "Sheet1" "G4")
    Assert.Equal(50.0, numberAt roundTripped "Sheet1" "H4")

[<Fact>]
let ``example: pivot table sourced from another sheet`` () =
    let sourceSheet =
        sheet
            "Data"
            [ row [ cell (Text "Category"); cell (Text "Amount") ]
              row [ cell (Text "A"); cell (Number 3.0) ]
              row [ cell (Text "B"); cell (Number 7.0) ]
              row [ cell (Text "A"); cell (Number 4.0) ] ]

    let reportSheet =
        sheet
            "Report"
            [ row [ cell (Text "Pivot table below:") ]
              EmbeddedPivotTable
                  { SourceSheet = Some "Data"
                    SourceTopLeft = CellRef.ofA1 "A1"
                    SourceBottomRight = CellRef.ofA1 "B4"
                    RowField = "Category"
                    ColumnField = None
                    ValueField = "Amount"
                    Aggregation = PivotCount
                    ValueCaption = None
                    TopLeftAnchor = CellRef.ofA1 "A3" } ]

    let scenarioName = "PivotTableAcrossSheets"
    verifyScenario scenarioName (workbook [ sourceSheet; reportSheet ])

    let roundTripped = Workbook.load (Path.Combine(examplesDir, scenarioName, "output.xlsx"))
    // A3 Category | B3 Count of Amount
    // A4 A        | B4 2
    // A5 B        | B5 1
    // A6 Grand Total | B6 3
    Assert.Equal(2.0, numberAt roundTripped "Report" "B4")
    Assert.Equal(1.0, numberAt roundTripped "Report" "B5")
    Assert.Equal(3.0, numberAt roundTripped "Report" "B6")

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

    // Almost every scenario's saved file is "output.xlsx" - the one exception is a VBA
    // macro scenario, which needs the ".xlsm" extension real Excel expects for
    // macro-enabled content (see `verifyScenarioNamed`), so this is discovered from disk
    // rather than hard-coded.
    let outputPath =
        Directory.GetFiles(dir) |> Array.find (fun f -> not (f.EndsWith("script.fsx")))

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
