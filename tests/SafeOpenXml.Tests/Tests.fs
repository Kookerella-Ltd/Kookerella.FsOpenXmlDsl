module SafeOpenXml.Tests

open System
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

let private headerStyle: CellStyle =
    { CellStyle.Default with
        Font = Some { FontStyle.Default with Bold = true; Color = Some Color.white }
        Fill = Some { Color = Rgb(68uy, 84uy, 106uy) }
        Border =
            Some
                { BorderStyle.None with
                    Bottom = Some { Style = Thin; Color = Some Color.black } } }

let private buildSampleWorkbook () : Workbook =
    let dateStyle = { CellStyle.Default with NumberFormat = Some ShortDate }
    let twoDecimal = { CellStyle.Default with NumberFormat = Some TwoDecimal }

    let data =
        sheet
            "Data"
            [ row [ cell (Text "Name", style = headerStyle)
                    cell (Text "Amount", style = headerStyle)
                    cell (Text "Purchased", style = headerStyle) ]
              row [ cell (Text "Widgets")
                    cell (Number 42.5, style = twoDecimal)
                    cell (Date(DateTime(2026, 1, 15)), style = dateStyle) ]
              row [ cell (Text "In stock"); cell (Boolean true) ]
              row [ cell (Text "Total"); cell (Formula("SUM(B2:B2)", Some 42.5)) ]
              ColumnWidth(0, 20.0)
              RowHeight(0, 20.0)
              Merge(CellRef.ofA1 "A5", CellRef.ofA1 "C5")
              Freeze(1, 0) ]

    workbook [ data ]

[<Fact>]
let ``writer produces a schema-valid workbook`` () =
    let wb = buildSampleWorkbook ()
    use stream = new MemoryStream()
    Workbook.saveToStream stream wb
    stream.Position <- 0L

    use document = SpreadsheetDocument.Open(stream, false)
    let validator = OpenXmlValidator()
    let errors = validator.Validate(document) |> List.ofSeq

    Assert.True(
        errors.IsEmpty,
        String.Join("\n", errors |> Seq.map (fun e -> sprintf "%s: %s" e.Path.XPath e.Description))
    )

[<Fact>]
let ``round trips cell values and styles through a real xlsx`` () =
    let original = buildSampleWorkbook ()
    use stream = new MemoryStream()
    Workbook.saveToStream stream original
    stream.Position <- 0L
    let roundTripped = Workbook.loadFromStream stream

    Assert.Equal(1, roundTripped.Sheets.Length)
    let sheet = roundTripped.Sheets.[0]
    Assert.Equal("Data", sheet.Name)

    let originalCells = original.Sheets.[0].Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col)
    let roundTrippedCells = sheet.Cells |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col)

    Assert.Equal(originalCells.Length, roundTrippedCells.Length)

    List.zip originalCells roundTrippedCells
    |> List.iter (fun (expected, actual) ->
        Assert.Equal<CellRef>(expected.Ref, actual.Ref)
        Assert.Equal<CellValue>(expected.Value, actual.Value)
        Assert.Equal<CellStyle option>(expected.Style, actual.Style))

    Assert.Equal<Map<int, ColumnProps>>(original.Sheets.[0].ColumnProps, sheet.ColumnProps)
    Assert.Equal<Map<int, RowProps>>(original.Sheets.[0].RowProps, sheet.RowProps)
    Assert.Equal<MergedRange list>(original.Sheets.[0].MergedRanges, sheet.MergedRanges)
    Assert.Equal<FreezePane option>(original.Sheets.[0].FreezePane, sheet.FreezePane)

let private redFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(255uy, 199uy, 206uy) } }
let private greenFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(198uy, 239uy, 206uy) } }

let private buildValidationSampleWorkbook () : Workbook =
    let data =
        sheet
            "Rules"
            [ row [ cell (Number 50.0); cell (Number 150.0); cell (Text "Apple"); cell (Number 5.0) ]
              row [ cell (Number 200.0); cell (Number 20.0); cell (Text "Apple"); cell (Number 10.0) ]
              conditionalFormat (CellRef.ofA1 "A1", CellRef.ofA1 "A2", CellValueRule(GreaterThan, "100", None, redFillStyle))
              conditionalFormat (CellRef.ofA1 "B1", CellRef.ofA1 "B2", FormulaRule("B1>A1", greenFillStyle))
              conditionalFormat (CellRef.ofA1 "D1", CellRef.ofA1 "D2", ColorScale2(Color.white, Color.red))
              conditionalFormat (CellRef.ofA1 "D3", CellRef.ofA1 "D4", ColorScale3(Color.red, Color.yellow, Color.green))
              conditionalFormat (CellRef.ofA1 "E1", CellRef.ofA1 "E2", DataBarRule Color.blue)
              conditionalFormat (CellRef.ofA1 "C1", CellRef.ofA1 "C2", DuplicateValuesRule redFillStyle)
              conditionalFormat (CellRef.ofA1 "C3", CellRef.ofA1 "C4", UniqueValuesRule greenFillStyle)
              dataValidation (CellRef.ofA1 "F1", CellRef.ofA1 "F1", ListValidation [ "Yes"; "No"; "Maybe" ])
              dataValidation (CellRef.ofA1 "F2", CellRef.ofA1 "F2", ListFromRangeValidation(CellRef.ofA1 "A1", CellRef.ofA1 "A2"))
              dataValidation (
                  CellRef.ofA1 "F3",
                  CellRef.ofA1 "F3",
                  WholeNumberValidation(GreaterThan, "0", None),
                  errorTitle = "Bad input",
                  errorMessage = "Must be positive"
              )
              dataValidation (CellRef.ofA1 "F4", CellRef.ofA1 "F4", DecimalValidation(Between, "0", Some "1"))
              dataValidation (CellRef.ofA1 "F5", CellRef.ofA1 "F5", TextLengthValidation(LessThanOrEqual, "10", None))
              dataValidation (
                  CellRef.ofA1 "F6",
                  CellRef.ofA1 "F6",
                  CustomValidation("ISNUMBER(F6)"),
                  allowBlank = false,
                  inputTitle = "Note",
                  inputMessage = "Enter a number"
              ) ]

    workbook [ data ]

[<Fact>]
let ``conditional formatting and data validation produce a schema-valid workbook`` () =
    let wb = buildValidationSampleWorkbook ()
    use stream = new MemoryStream()
    Workbook.saveToStream stream wb
    stream.Position <- 0L

    use document = SpreadsheetDocument.Open(stream, false)
    let validator = OpenXmlValidator()
    let errors = validator.Validate(document) |> List.ofSeq

    Assert.True(
        errors.IsEmpty,
        String.Join("\n", errors |> Seq.map (fun e -> sprintf "%s: %s" e.Path.XPath e.Description))
    )

[<Fact>]
let ``conditional formatting and data validation round trip`` () =
    let original = buildValidationSampleWorkbook ()
    use stream = new MemoryStream()
    Workbook.saveToStream stream original
    stream.Position <- 0L
    let roundTripped = Workbook.loadFromStream stream

    Assert.Equal<ConditionalFormatEntry list>(original.Sheets.[0].ConditionalFormats, roundTripped.Sheets.[0].ConditionalFormats)
    Assert.Equal<DataValidationEntry list>(original.Sheets.[0].DataValidations, roundTripped.Sheets.[0].DataValidations)

[<Fact>]
let ``multiple sheets round trip`` () =
    let wb =
        workbook
            [ sheetOfCells "First" [ cellA1 "A1" (Text "one") ]
              sheetOfCells "Second" [ cellA1 "A1" (Number 2.0) ] ]

    use stream = new MemoryStream()
    Workbook.saveToStream stream wb
    stream.Position <- 0L
    let roundTripped = Workbook.loadFromStream stream

    Assert.Equal<string list>([ "First"; "Second" ], roundTripped.Sheets |> List.map (fun s -> s.Name))
