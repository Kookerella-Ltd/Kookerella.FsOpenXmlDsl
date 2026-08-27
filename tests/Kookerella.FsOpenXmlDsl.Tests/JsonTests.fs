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

/// The canonical "1x1 transparent GIF" - same fixture `XmlTests.fs`/`Tests.fs` use, for the
/// same reason: the smallest possible valid image file, well-known and trustworthy as a
/// test fixture.
let private onePixelGif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7")

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

[<Fact>]
let ``Json round trip: Comments`` () =
    let sheet =
        { emptySheet "Sheet1" with
            Comments =
                [ { Cell = CellRef.create 0 0; Author = "Alex"; Text = "Check this figure" }
                  { Cell = CellRef.create 1 0; Author = ""; Text = "Unnamed author" } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: Hyperlinks, external and internal`` () =
    let sheet =
        { emptySheet "Sheet1" with
            Hyperlinks =
                [ { TopLeft = CellRef.create 0 0
                    BottomRight = CellRef.create 0 0
                    Target = ExternalHyperlink "https://example.com"
                    Tooltip = Some "Visit site"
                    Display = None }
                  { TopLeft = CellRef.create 1 0
                    BottomRight = CellRef.create 2 1
                    Target = InternalHyperlink "Sheet1!A1"
                    Tooltip = None
                    Display = Some "Go to top" } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: SheetProtection`` () =
    let protection =
        { SheetProtection.Default with
            Password = Some "hunter2"
            FormatCells = Some true
            Sort = Some true
            AutoFilter = Some true }

    let sheet = { emptySheet "Sheet1" with Protection = Some protection }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: SheetProtection default has no optional flags set`` () =
    let sheet = { emptySheet "Sheet1" with Protection = Some SheetProtection.Default }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: WorkbookProtection`` () =
    let original =
        { emptyWorkbook [ emptySheet "Sheet1" ] with
            Protection = Some { WorkbookProtection.Default with Password = Some "hunter2"; LockStructure = Some true } }

    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PageSetup, landscape with custom margins and named paper size`` () =
    let pageSetup =
        { PageSetup.Default with
            Orientation = Landscape
            PaperSize = Some A4
            Margins = { Left = 0.5; Right = 0.5; Top = 1.0; Bottom = 1.0; Header = 0.2; Footer = 0.2 } }

    let sheet = { emptySheet "Sheet1" with PageSetup = Some pageSetup }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PageSetup, fit-to-page scaling and OtherPaperSize`` () =
    let pageSetup =
        { PageSetup.Default with
            Scaling = Some(FitToPage(1, 0))
            PaperSize = Some(OtherPaperSize 9) }

    let sheet = { emptySheet "Sheet1" with PageSetup = Some pageSetup }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PageSetup, print area and header/footer variants`` () =
    let pageSetup =
        { PageSetup.Default with
            Scaling = Some(ScalePercent 80)
            PrintArea = [ CellRef.create 0 0, CellRef.create 9 3; CellRef.create 20 0, CellRef.create 25 1 ]
            Header = Some "&C&\"Arial,Bold\"Report"
            Footer = Some "&LPage &P of &N"
            EvenHeader = Some "Even page"
            EvenFooter = Some "Even footer"
            FirstHeader = Some "Cover page"
            FirstFooter = Some "Cover footer" }

    let sheet = { emptySheet "Sheet1" with PageSetup = Some pageSetup }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PageSetup.Default omits unset optional fields`` () =
    let sheet = { emptySheet "Sheet1" with PageSetup = Some PageSetup.Default }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: Images`` () =
    let sheet =
        { emptySheet "Sheet1" with
            Images =
                [ { Data = onePixelGif
                    Format = Gif
                    TopLeftAnchor = CellRef.create 0 0
                    BottomRightAnchor = CellRef.create 5 3 } ] }

    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook

    Assert.Equal<Workbook>(original, roundTripped)
    Assert.Equal<byte[]>(onePixelGif, (roundTripped.Sheets |> List.exactlyOne).Images.[0].Data)

[<Fact>]
let ``Json round trip: Table with default style`` () =
    let table =
        { TopLeft = CellRef.create 0 0
          BottomRight = CellRef.create 2 2
          Name = "SalesTable"
          Columns = [ { Name = "Item"; CalculatedFormula = None }; { Name = "Amount"; CalculatedFormula = None } ]
          Style = TableStyle.Default }

    let sheet = { emptySheet "Sheet1" with Tables = [ table ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: Table with calculated column and custom style`` () =
    let table =
        { TopLeft = CellRef.create 0 0
          BottomRight = CellRef.create 3 1
          Name = "Calc"
          Columns =
            [ { Name = "Qty"; CalculatedFormula = None }
              { Name = "Doubled"; CalculatedFormula = Some "Calc[Qty]*2" } ]
          Style =
            { Name = Some "TableStyleLight9"
              ShowFirstColumn = true
              ShowLastColumn = true
              ShowRowStripes = false
              ShowColumnStripes = true } }

    let sheet = { emptySheet "Sheet1" with Tables = [ table ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: SparklineGroup, line with high/low markers`` () =
    let group =
        { Style = { SparklineStyle.Default with ShowHigh = true; ShowLow = true }
          Sparklines =
            [ { Cell = CellRef.ofA1 "F1"; DataTopLeft = CellRef.ofA1 "B1"; DataBottomRight = CellRef.ofA1 "E1" }
              { Cell = CellRef.ofA1 "F2"; DataTopLeft = CellRef.ofA1 "B2"; DataBottomRight = CellRef.ofA1 "E2" } ] }

    let sheet = { emptySheet "Sheet1" with SparklineGroups = [ group ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: SparklineGroup, column with custom color and negative highlighting`` () =
    let group =
        { Style =
            { SparklineStyle.Default with
                Type = SparklineType.Column
                Color = Some(Rgb(0uy, 112uy, 192uy))
                LineWeight = Some 1.5
                ShowNegative = true }
          Sparklines = [ { Cell = CellRef.ofA1 "E1"; DataTopLeft = CellRef.ofA1 "A1"; DataBottomRight = CellRef.ofA1 "D1" } ] }

    let sheet = { emptySheet "Sheet1" with SparklineGroups = [ group ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: Chart, column with title, legend, and two series`` () =
    let chart =
        { Type = ChartColumn
          Title = Some "Sales by Quarter"
          CategoriesTopLeft = CellRef.ofA1 "A2"
          CategoriesBottomRight = CellRef.ofA1 "A4"
          Series =
            [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B4" }
              { Name = CellRef.ofA1 "C1"; ValuesTopLeft = CellRef.ofA1 "C2"; ValuesBottomRight = CellRef.ofA1 "C4" } ]
          ShowLegend = true
          TopLeftAnchor = CellRef.ofA1 "E1"
          BottomRightAnchor = CellRef.ofA1 "L15" }

    let sheet = { emptySheet "Sheet1" with Charts = [ chart ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: Chart, bar with no title and one series`` () =
    let chart =
        { Type = ChartBar
          Title = None
          CategoriesTopLeft = CellRef.ofA1 "A2"
          CategoriesBottomRight = CellRef.ofA1 "A3"
          Series = [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B3" } ]
          ShowLegend = false
          TopLeftAnchor = CellRef.ofA1 "D1"
          BottomRightAnchor = CellRef.ofA1 "K12" }

    let sheet = { emptySheet "Sheet1" with Charts = [ chart ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PivotTable, row field only`` () =
    let pivot =
        { SourceSheet = None
          SourceTopLeft = CellRef.ofA1 "A1"
          SourceBottomRight = CellRef.ofA1 "B5"
          RowField = "Region"
          ColumnField = None
          ValueField = "Sales"
          Aggregation = PivotSum
          ValueCaption = None
          TopLeftAnchor = CellRef.ofA1 "D1" }

    let sheet = { emptySheet "Sheet1" with PivotTables = [ pivot ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: PivotTable, row and column fields, cross-sheet source, and value caption`` () =
    let pivot =
        { SourceSheet = Some "Data"
          SourceTopLeft = CellRef.ofA1 "A1"
          SourceBottomRight = CellRef.ofA1 "C9"
          RowField = "Region"
          ColumnField = Some "Quarter"
          ValueField = "Sales"
          Aggregation = PivotAverage
          ValueCaption = Some "Avg Sales"
          TopLeftAnchor = CellRef.ofA1 "F1" }

    let sheet = { emptySheet "Sheet1" with PivotTables = [ pivot ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

let private redFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(255uy, 199uy, 206uy) } }
let private greenFillStyle = { CellStyle.Default with Fill = Some { Color = Rgb(198uy, 239uy, 206uy) } }

[<Fact>]
let ``Json round trip: ConditionalFormat, CellValueRule`` () =
    let entry =
        { TopLeft = CellRef.ofA1 "A1"
          BottomRight = CellRef.ofA1 "A3"
          Rule = CellValueRule(GreaterThan, "100", None, redFillStyle) }

    let sheet = { emptySheet "Sheet1" with ConditionalFormats = [ entry ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: ConditionalFormat, FormulaRule`` () =
    let entry =
        { TopLeft = CellRef.ofA1 "A1"
          BottomRight = CellRef.ofA1 "A2"
          Rule = FormulaRule("A1>B1", greenFillStyle) }

    let sheet = { emptySheet "Sheet1" with ConditionalFormats = [ entry ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: ConditionalFormat, ColorScale2 and ColorScale3`` () =
    let entries =
        [ { TopLeft = CellRef.ofA1 "A1"; BottomRight = CellRef.ofA1 "A3"; Rule = ColorScale2(Color.white, Color.red) }
          { TopLeft = CellRef.ofA1 "B1"
            BottomRight = CellRef.ofA1 "B3"
            Rule = ColorScale3(Color.red, Color.yellow, Color.green) } ]

    let sheet = { emptySheet "Sheet1" with ConditionalFormats = entries }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: ConditionalFormat, DataBarRule, DuplicateValuesRule, and UniqueValuesRule`` () =
    let entries =
        [ { TopLeft = CellRef.ofA1 "A1"; BottomRight = CellRef.ofA1 "A3"; Rule = DataBarRule Color.blue }
          { TopLeft = CellRef.ofA1 "B1"; BottomRight = CellRef.ofA1 "B3"; Rule = DuplicateValuesRule redFillStyle }
          { TopLeft = CellRef.ofA1 "C1"; BottomRight = CellRef.ofA1 "C3"; Rule = UniqueValuesRule greenFillStyle } ]

    let sheet = { emptySheet "Sheet1" with ConditionalFormats = entries }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: DataValidation, ListValidation and ListFromRangeValidation`` () =
    let entries =
        [ { TopLeft = CellRef.ofA1 "A2"
            BottomRight = CellRef.ofA1 "A2"
            Kind = ListValidation [ "Small"; "Medium"; "Large" ]
            Alert = ValidationAlert.Default }
          { TopLeft = CellRef.ofA1 "B2"
            BottomRight = CellRef.ofA1 "B2"
            Kind = ListFromRangeValidation(CellRef.ofA1 "A1", CellRef.ofA1 "C1")
            Alert = ValidationAlert.Default } ]

    let sheet = { emptySheet "Sheet1" with DataValidations = entries }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: DataValidation, WholeNumberValidation with error alert`` () =
    let entry =
        { TopLeft = CellRef.ofA1 "A2"
          BottomRight = CellRef.ofA1 "A2"
          Kind = WholeNumberValidation(GreaterThan, "0", None)
          Alert =
            { ValidationAlert.Default with
                ErrorTitle = Some "Invalid quantity"
                ErrorMessage = Some "Quantity must be a positive whole number." } }

    let sheet = { emptySheet "Sheet1" with DataValidations = [ entry ] }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: DataValidation, DecimalValidation and TextLengthValidation`` () =
    let entries =
        [ { TopLeft = CellRef.ofA1 "A2"
            BottomRight = CellRef.ofA1 "A2"
            Kind = DecimalValidation(Between, "0", Some "1")
            Alert = ValidationAlert.Default }
          { TopLeft = CellRef.ofA1 "B2"
            BottomRight = CellRef.ofA1 "B2"
            Kind = TextLengthValidation(LessThanOrEqual, "10", None)
            Alert = ValidationAlert.Default } ]

    let sheet = { emptySheet "Sheet1" with DataValidations = entries }
    let original = emptyWorkbook [ sheet ]
    let roundTripped = original |> Json.ofWorkbook |> Json.toWorkbook
    Assert.Equal<Workbook>(original, roundTripped)

[<Fact>]
let ``Json round trip: DataValidation, CustomValidation with allowBlank false, warning style, and input prompt`` () =
    let entry =
        { TopLeft = CellRef.ofA1 "A2"
          BottomRight = CellRef.ofA1 "A2"
          Kind = CustomValidation "ISNUMBER(A2)"
          Alert =
            { AllowBlank = false
              ErrorStyle = Warning
              ErrorTitle = None
              ErrorMessage = None
              InputTitle = Some "Note"
              InputMessage = Some "Enter a numeric value." } }

    let sheet = { emptySheet "Sheet1" with DataValidations = [ entry ] }
    let original = emptyWorkbook [ sheet ]
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
                  { TopLeft = CellRef.ofA1 "A1"; BottomRight = CellRef.ofA1 "B1" } ]
            Comments =
                [ { Cell = CellRef.ofA1 "B2"; Author = "Alex"; Text = "Second" }
                  { Cell = CellRef.ofA1 "A1"; Author = "Alex"; Text = "First" } ] }

    let sheetB =
        { sheetA with
            Cells = sheetA.Cells |> List.rev
            MergedRanges = sheetA.MergedRanges |> List.rev
            Comments = sheetA.Comments |> List.rev }

    let wbA =
        { emptyWorkbook [ sheetA ] with
            DefinedNames =
                [ { Name = "Zeta"; Formula = "1"; Scope = WorkbookScope; Hidden = false }
                  { Name = "Alpha"; Formula = "2"; Scope = WorkbookScope; Hidden = false } ] }

    let wbB = { wbA with DefinedNames = wbA.DefinedNames |> List.rev; Sheets = [ sheetB ] }

    Assert.Equal((Json.ofWorkbook wbA).ToJsonString(), (Json.ofWorkbook wbB).ToJsonString())
