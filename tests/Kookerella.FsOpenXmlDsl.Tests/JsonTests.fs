module Kookerella.FsOpenXmlDsl.JsonTests

open System
open Xunit
open Kookerella.FsOpenXmlDsl

/// Mirrors `XmlTests.fs`'s `emptySheet`/`emptyWorkbook` - a `Worksheet`/`Workbook` with
/// every field this pass of `Json` doesn't yet model already at the "empty" value
/// `Json.toWorkbook` always produces for it, so a round trip is a meaningful check only on
/// the fields that are modeled so far.
let private emptySheet name =
    { Name = name
      Cells = []
      ColumnProps = Map.empty
      RowProps = Map.empty
      MergedRanges = []
      FreezePane = None
      AutoFilter = None
      Protection = None
      ConditionalFormats = []
      DataValidations = []
      Hyperlinks = []
      Comments = []
      PageSetup = None
      Tables = []
      SparklineGroups = []
      Charts = []
      Images = []
      PivotTables = [] }

let private emptyWorkbook sheets =
    { Sheets = sheets
      DefinedNames = []
      Protection = None
      VbaProject = None }

[<Fact>]
let ``Json round trip: every CellValue kind`` () =
    let sheet =
        { emptySheet "Sheet1" with
            Cells =
                [ { Ref = CellRef.create 0 0; Value = Text "Hello"; Style = None }
                  { Ref = CellRef.create 0 1; Value = Number 42.5; Style = None }
                  { Ref = CellRef.create 0 2; Value = Boolean true; Style = None }
                  { Ref = CellRef.create 0 3
                    Value = Date(DateTime(2024, 3, 15, 0, 0, 0, DateTimeKind.Utc))
                    Style = None }
                  { Ref = CellRef.create 0 4; Value = Formula("SUM(A1:B1)", Some 42.5); Style = None }
                  { Ref = CellRef.create 0 5; Value = Formula("A1", None); Style = None } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: full CellStyle`` () =
    let style =
        { CellStyle.Default with
            Font =
                Some
                    { FontStyle.Default with
                        Name = Some "Calibri"
                        Size = Some 11.0
                        Bold = true
                        Italic = true
                        Underline = true
                        Strikethrough = true
                        Color = Some(Rgb(12uy, 34uy, 56uy)) }
            Fill = Some { FillStyle.Color = Indexed 5 }
            Border =
                Some
                    { BorderStyle.None with
                        Left = Some { Style = Thin; Color = Some Color.black }
                        Right = Some { Style = Other "slantDashDot"; Color = None } }
            NumberFormat = Some Currency
            Alignment =
                Some
                    { Horizontal = Some AlignCenter
                      Vertical = Some AlignMiddle
                      WrapText = true }
            Protection = Some { Locked = false; Hidden = true } }

    let customStyle = { CellStyle.Default with NumberFormat = Some(Custom "0.000%") }

    let themeColorStyle =
        { CellStyle.Default with
            Font = Some { FontStyle.Default with Color = Some(Theme(2, Some 0.5)) } }

    let sheet =
        { emptySheet "Sheet1" with
            Cells =
                [ { Ref = CellRef.create 0 0; Value = Number 1.0; Style = Some style }
                  { Ref = CellRef.create 1 0; Value = Number 2.0; Style = Some customStyle }
                  { Ref = CellRef.create 2 0
                    Value = Number 3.0
                    Style = Some themeColorStyle } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: merged ranges, freeze pane, autofilter, column and row sizing`` () =
    let sheet =
        { emptySheet "Sheet1" with
            Cells = [ { Ref = CellRef.create 0 0; Value = Text "Header"; Style = None } ]
            MergedRanges =
                [ { TopLeft = CellRef.create 0 0; BottomRight = CellRef.create 0 2 }
                  { TopLeft = CellRef.create 3 0; BottomRight = CellRef.create 4 1 } ]
            FreezePane = Some { Rows = 1; Columns = 0 }
            AutoFilter = Some { TopLeft = CellRef.create 0 0; BottomRight = CellRef.create 10 3 }
            ColumnProps = Map.ofList [ 0, { Width = Some 20.0 }; 2, { Width = None } ]
            RowProps = Map.ofList [ 0, { Height = Some 30.0 } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: multiple sheets and an empty workbook`` () =
    let original = emptyWorkbook [ emptySheet "First"; emptySheet "Second" ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: VbaProject`` () =
    let bytes = IO.File.ReadAllBytes(IO.Path.Combine(__SOURCE_DIRECTORY__, "Assets", "sample.vbaProject.bin"))
    let original = { emptyWorkbook [ emptySheet "Sheet1" ] with VbaProject = Some bytes }
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)
    Assert.Equal<byte[]>(bytes, roundTripped.VbaProject.Value)

[<Fact>]
let ``Json round trip: no VbaProject stays None`` () =
    let original = emptyWorkbook [ emptySheet "Sheet1" ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.True(roundTripped.VbaProject.IsNone)

[<Fact>]
let ``Json round trip: DefinedNames, workbook and sheet scoped`` () =
    let original =
        { emptyWorkbook [ emptySheet "Sheet1" ] with
            DefinedNames =
                [ { Name = "LocalTotal"
                    Formula = "Sheet1!$A$2"
                    Scope = SheetScope "Sheet1"
                    Hidden = true }
                  { Name = "TaxRate"; Formula = "0.075"; Scope = WorkbookScope; Hidden = false } ] }

    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

/// Mirrors `` Xml.ofWorkbook produces deterministic, input-order-independent output `` -
/// the same property matters for JSON: two `Workbook` values with the same content but
/// differently-ordered lists must render to identical JSON text, so a re-generated
/// `workbook.json` committed to source control doesn't show a spurious diff.
[<Fact>]
let ``Json.ofWorkbook produces deterministic, input-order-independent output`` () =
    let sheetA =
        { emptySheet "Sheet1" with
            Cells =
                [ { Ref = CellRef.ofA1 "B2"; Value = Number 2.0; Style = None }
                  { Ref = CellRef.ofA1 "A1"; Value = Number 1.0; Style = None }
                  { Ref = CellRef.ofA1 "A2"; Value = Number 3.0; Style = None } ]
            MergedRanges =
                [ { TopLeft = CellRef.ofA1 "C1"; BottomRight = CellRef.ofA1 "D1" }
                  { TopLeft = CellRef.ofA1 "A1"; BottomRight = CellRef.ofA1 "B1" } ] }

    let sheetB =
        { sheetA with
            Cells = sheetA.Cells |> List.rev
            MergedRanges = sheetA.MergedRanges |> List.rev }

    let wbA =
        { emptyWorkbook [ sheetA ] with
            DefinedNames =
                [ { Name = "Zeta"; Formula = "1"; Scope = WorkbookScope; Hidden = false }
                  { Name = "Alpha"; Formula = "2"; Scope = WorkbookScope; Hidden = false } ] }

    let wbB = { wbA with DefinedNames = wbA.DefinedNames |> List.rev; Sheets = [ sheetB ] }

    Assert.Equal((Json.ofWorkbook wbA).ToJsonString(), (Json.ofWorkbook wbB).ToJsonString())
