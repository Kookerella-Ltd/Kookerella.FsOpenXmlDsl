using System.Diagnostics;
using System.Runtime.CompilerServices;
using Kookerella.CsOpenXmlDsl;
using Xunit;
using static Kookerella.CsOpenXmlDsl.Tests.TestHelpers;

namespace Kookerella.CsOpenXmlDsl.Tests;

/// <summary>
/// Each scenario below is a self-contained demonstration of one feature - mirroring the F#
/// test suite's own <c>Examples/</c> convention. Running it writes the workbook it builds to
/// <c>Examples/&lt;scenario name&gt;/output.xlsx</c> (checked into the repo, so you can open
/// any single feature in Excel without re-running anything) plus an
/// <c>Examples/&lt;scenario name&gt;/script.cs</c> generated via <see cref="CsCodeGen"/> -
/// a real, directly-runnable <c>dotnet run script.cs</c> file, not just a fragment.
/// <para>
/// Unlike the F# suite's own generated <c>script.fsx</c> files (which embed an absolute,
/// machine-specific <c>#r</c> path resolved via reflection at test-run time), the <c>#:
/// project</c> line used here is a plain path relative to the script's own folder - so the
/// committed <c>script.cs</c> is identical across machines and CI runs, and genuinely
/// portable: clone the repo, <c>cd</c> into any scenario folder, and <c>dotnet run
/// script.cs</c> just works.
/// </para>
/// </summary>
public class ExampleTests
{
    /// <summary>The C# analogue of F#'s <c>__SOURCE_DIRECTORY__</c> - captures this file's
    /// own compile-time path so <c>Examples/</c> resolves correctly regardless of build
    /// configuration or working directory.</summary>
    private static string ExamplesDir([CallerFilePath] string sourceFile = "") =>
        Path.Combine(Path.GetDirectoryName(sourceFile)!, "Examples");

    /// <summary>4 directories up from <c>Examples/&lt;name&gt;/script.cs</c> reaches the
    /// repo root (Examples -&gt; Kookerella.CsOpenXmlDsl.Tests -&gt; tests -&gt; root).</summary>
    private const string ProjectReferenceLine = "#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj";

    /// <summary>Saves <paramref name="workbook"/> to <c>Examples/&lt;name&gt;/&lt;fileName&gt;</c>,
    /// asserts the file is schema-valid, writes the scenario's <c>script.cs</c> as a side
    /// effect, and returns the loaded-back workbook so the caller can assert whatever facts
    /// matter for that specific feature - the same two checks (schema-valid, round-trips)
    /// every scenario needs, factored out so each one only states what it's building.</summary>
    private static Workbook VerifyScenario(string name, Workbook workbook, string fileName = "output.xlsx")
    {
        var dir = Path.Combine(ExamplesDir(), name);
        Directory.CreateDirectory(dir);

        var outputPath = Path.Combine(dir, fileName);
        WorkbookIO.Save(workbook, outputPath);
        AssertSchemaValid(outputPath);

        var script = CsCodeGen.Generate([ProjectReferenceLine], fileName, workbook);
        File.WriteAllText(Path.Combine(dir, "script.cs"), script);

        return WorkbookIO.Load(outputPath);
    }

    [Fact]
    public void TextNumberBooleanDateCells()
    {
        var sheet = Sheet.Create(
            "Sheet1",
            Row.Of(Cell.Text("Item"), Cell.Text("Qty"), Cell.Text("Price"), Cell.Text("Total")),
            Row.Of(Cell.Text("Widgets"), Cell.Number(4), Cell.Number(2.5), Cell.Formula("B2*C2", 10.0)),
            Row.Of(Cell.Boolean(true), Cell.Date(new DateTime(2026, 1, 1))));

        var loaded = VerifyScenario(nameof(TextNumberBooleanDateCells), Workbook.Create(sheet));
        var loadedSheet = Assert.Single(loaded.Sheets);
        Assert.Equal("Sheet1", loadedSheet.Name);

        var byPosition = loadedSheet.Rows.OrderBy(r => r.Index).ToArray();
        Assert.Equal(3, byPosition.Length);

        var row2 = byPosition[1].Cells.OrderBy(c => c.Column).ToArray();
        Assert.Equal("Widgets", Assert.IsType<CellValue.Text>(row2[0].Value).Value);
        Assert.Equal(4.0, Assert.IsType<CellValue.Number>(row2[1].Value).Value);
        Assert.Equal(2.5, Assert.IsType<CellValue.Number>(row2[2].Value).Value);
        var formula = Assert.IsType<CellValue.Formula>(row2[3].Value);
        Assert.Equal("B2*C2", formula.Expression);
        Assert.Equal(10.0, formula.CachedValue);

        var row3 = byPosition[2].Cells.OrderBy(c => c.Column).ToArray();
        Assert.True(Assert.IsType<CellValue.Boolean>(row3[0].Value).Value);
        Assert.Equal(new DateTime(2026, 1, 1), Assert.IsType<CellValue.Date>(row3[1].Value).Value);
    }

    [Fact]
    public void SparseRowAndColumnJumps()
    {
        // Row 0 implicit, row 4 explicit (gap at 1-3), column 0 implicit then column 2 explicit.
        var sheet = Sheet.Create(
            "Sheet1",
            Row.Of(Cell.Text("A1"), Cell.Text("B1")),
            Row.Of(Cell.Text("A5"), Cell.Text("C5").AtColumn(2)).AtIndex(4));

        var loaded = VerifyScenario(nameof(SparseRowAndColumnJumps), Workbook.Create(sheet));
        var loadedSheet = Assert.Single(loaded.Sheets);

        var allCells = loadedSheet.Rows.SelectMany(r => r.Cells.Select(c => (Row: r.Index, Col: c.Column, c.Value)));
        Assert.Contains(allCells, c => c.Row == 0 && c.Col == 0 && ((CellValue.Text)c.Value!).Value == "A1");
        Assert.Contains(allCells, c => c.Row == 0 && c.Col == 1 && ((CellValue.Text)c.Value!).Value == "B1");
        Assert.Contains(allCells, c => c.Row == 4 && c.Col == 0 && ((CellValue.Text)c.Value!).Value == "A5");
        Assert.Contains(allCells, c => c.Row == 4 && c.Col == 2 && ((CellValue.Text)c.Value!).Value == "C5");
    }

    [Fact]
    public void StyledHeaderRow()
    {
        var style = CellStyle.Default
            .AsBold()
            .AsItalic()
            .WithFontColor(RgbColor.White)
            .WithFillColor(new RgbColor(68, 84, 106))
            .WithBorder(CellBorder.None.WithAllSides(new BorderSide(BorderLineStyle.Thin)))
            .WithHorizontalAlignment(HorizontalCellAlignment.Center)
            .AsWrapText()
            .WithNumberFormat(NumberFormatKind.Currency);

        var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Number(42.5).WithStyle(style)));

        var loaded = VerifyScenario(nameof(StyledHeaderRow), Workbook.Create(sheet));
        var cell = loaded.Sheets.Single().Rows.Single().Cells.Single();
        Assert.NotNull(cell.Style);
        Assert.True(cell.Style!.Bold);
        Assert.True(cell.Style.Italic);
        Assert.Equal(RgbColor.White, cell.Style.FontColor);
        Assert.Equal(new RgbColor(68, 84, 106), cell.Style.FillColor);
        Assert.Equal(HorizontalCellAlignment.Center, cell.Style.HorizontalAlignment);
        Assert.True(cell.Style.WrapText);
        Assert.Equal(NumberFormatKind.Currency, cell.Style.NumberFormat);
        Assert.NotNull(cell.Style.Border);
        Assert.Equal(BorderLineStyle.Thin, cell.Style.Border!.Left!.Style);
        Assert.Equal(BorderLineStyle.Thin, cell.Style.Border.Right!.Style);
        Assert.Equal(BorderLineStyle.Thin, cell.Style.Border.Top!.Style);
        Assert.Equal(BorderLineStyle.Thin, cell.Style.Border.Bottom!.Style);
    }

    [Fact]
    public void NumberFormats()
    {
        var style = CellStyle.Default.WithCustomNumberFormat("0.000%");
        var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Number(0.5).WithStyle(style)));

        var loaded = VerifyScenario(nameof(NumberFormats), Workbook.Create(sheet));
        var cell = loaded.Sheets.Single().Rows.Single().Cells.Single();

        Assert.Null(cell.Style!.NumberFormat);
        Assert.Equal("0.000%", cell.Style.CustomNumberFormat);
    }

    [Fact]
    public void MultipleSheets()
    {
        var wb = Workbook
            .Create(Sheet.Create("First", Row.Of(Cell.Text("a"))))
            .AddSheet(Sheet.Create("Second", Row.Of(Cell.Text("b"))));

        var loaded = VerifyScenario(nameof(MultipleSheets), wb);
        Assert.Equal(["First", "Second"], loaded.Sheets.Select(s => s.Name));
    }

    [Fact]
    public void MergedCells()
    {
        var headerStyle = CellStyle.Default.AsBold();

        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Quarterly Report").WithStyle(headerStyle)),
                Row.Of(Cell.Text("Q1"), Cell.Text("Q2"), Cell.Text("Q3"), Cell.Text("Q4")))
            .WithMergedRanges(MergedRange.Of("A1", "D1"));

        var loaded = VerifyScenario(nameof(MergedCells), Workbook.Create(sheet));
        var loadedSheet = loaded.Sheets.Single();

        var mergedRange = Assert.Single(loadedSheet.MergedRanges);
        Assert.Equal(CellPosition.FromA1("A1"), mergedRange.TopLeft);
        Assert.Equal(CellPosition.FromA1("D1"), mergedRange.BottomRight);
    }

    [Fact]
    public void FrozenPanes()
    {
        var headerStyle = CellStyle.Default.AsBold();

        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Name").WithStyle(headerStyle), Cell.Text("Amount").WithStyle(headerStyle)),
                Row.Of(Cell.Text("Row 1"), Cell.Number(1)),
                Row.Of(Cell.Text("Row 2"), Cell.Number(2)))
            .WithFreezePane(1, 0);

        var loaded = VerifyScenario(nameof(FrozenPanes), Workbook.Create(sheet));
        var loadedSheet = loaded.Sheets.Single();

        Assert.NotNull(loadedSheet.FreezePane);
        Assert.Equal(1, loadedSheet.FreezePane!.Rows);
        Assert.Equal(0, loadedSheet.FreezePane.Columns);
    }

    [Fact]
    public void AutoFilter()
    {
        var headerStyle = CellStyle.Default.AsBold();

        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Name").WithStyle(headerStyle), Cell.Text("Amount").WithStyle(headerStyle), Cell.Text("Region").WithStyle(headerStyle)),
                Row.Of(Cell.Text("Widgets"), Cell.Number(42.5), Cell.Text("North")),
                Row.Of(Cell.Text("Gadgets"), Cell.Number(19.99), Cell.Text("South")))
            .WithAutoFilter(AutoFilterRange.Of("A1", "C3"));

        var loaded = VerifyScenario(nameof(AutoFilter), Workbook.Create(sheet));
        var loadedSheet = loaded.Sheets.Single();

        Assert.NotNull(loadedSheet.AutoFilter);
        Assert.Equal(CellPosition.FromA1("A1"), loadedSheet.AutoFilter!.TopLeft);
        Assert.Equal(CellPosition.FromA1("C3"), loadedSheet.AutoFilter.BottomRight);
    }

    [Fact]
    public void Table()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Item"), Cell.Text("Quantity")),
                Row.Of(Cell.Text("Widgets"), Cell.Number(12)),
                Row.Of(Cell.Text("Gadgets"), Cell.Number(5)))
            .WithTables(TableEntry.Of("A1", "B3", "Inventory", new TableColumn("Item"), new TableColumn("Quantity")));

        var loaded = VerifyScenario(nameof(Table), Workbook.Create(sheet));
        var table = Assert.Single(loaded.Sheets.Single().Tables);

        Assert.Equal("Inventory", table.Name);
        Assert.Equal(CellPosition.FromA1("A1"), table.TopLeft);
        Assert.Equal(CellPosition.FromA1("B3"), table.BottomRight);
        Assert.Equal(["Item", "Quantity"], table.Columns.Select(c => c.Name));
        Assert.All(table.Columns, c => Assert.Null(c.CalculatedFormula));
        Assert.Equal("TableStyleMedium2", table.Style.Name);
        Assert.True(table.Style.ShowRowStripes);
    }

    [Fact]
    public void TableWithCalculatedColumn()
    {
        var style = TableStyle.Default.WithName("TableStyleLight9").WithColumnStripes().WithoutRowStripes();

        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Quantity"), Cell.Text("Unit Price"), Cell.Text("Total")),
                Row.Of(Cell.Number(12), Cell.Number(2.5), Cell.Formula("[@Quantity]*[@[Unit Price]]", 30.0)),
                Row.Of(Cell.Number(5), Cell.Number(9), Cell.Formula("[@Quantity]*[@[Unit Price]]", 45.0)))
            .WithTables(
                TableEntry
                    .Of(
                        "A1",
                        "C3",
                        "Orders",
                        new TableColumn("Quantity"),
                        new TableColumn("Unit Price"),
                        new TableColumn("Total", "[@Quantity]*[@[Unit Price]]"))
                    .WithStyle(style));

        var loaded = VerifyScenario(nameof(TableWithCalculatedColumn), Workbook.Create(sheet));
        var table = Assert.Single(loaded.Sheets.Single().Tables);

        var total = table.Columns.Single(c => c.Name == "Total");
        Assert.Equal("[@Quantity]*[@[Unit Price]]", total.CalculatedFormula);
        Assert.Equal("TableStyleLight9", table.Style.Name);
        Assert.True(table.Style.ShowColumnStripes);
        Assert.False(table.Style.ShowRowStripes);
    }

    [Fact]
    public void ChartColumn()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Quarter"), Cell.Text("North"), Cell.Text("South")),
                Row.Of(Cell.Text("Q1"), Cell.Number(12), Cell.Number(9)),
                Row.Of(Cell.Text("Q2"), Cell.Number(15), Cell.Number(11)),
                Row.Of(Cell.Text("Q3"), Cell.Number(9), Cell.Number(14)))
            .AddChart(
                ChartEntry
                    .Of(
                        ChartType.Column,
                        "A2",
                        "A4",
                        "E1",
                        "L15",
                        ChartSeries.Of("B1", "B2", "B4"),
                        ChartSeries.Of("C1", "C2", "C4"))
                    .WithTitle("Sales by Quarter")
                    .WithLegend());

        var loaded = VerifyScenario(nameof(ChartColumn), Workbook.Create(sheet));
        var chart = Assert.Single(loaded.Sheets.Single().Charts);

        Assert.Equal(ChartType.Column, chart.Type);
        Assert.Equal("Sales by Quarter", chart.Title);
        Assert.True(chart.ShowLegend);
        Assert.Equal(CellPosition.FromA1("A2"), chart.CategoriesTopLeft);
        Assert.Equal(CellPosition.FromA1("A4"), chart.CategoriesBottomRight);
        Assert.Equal(CellPosition.FromA1("E1"), chart.TopLeftAnchor);
        Assert.Equal(CellPosition.FromA1("L15"), chart.BottomRightAnchor);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal(CellPosition.FromA1("B1"), chart.Series[0].Name);
        Assert.Equal(CellPosition.FromA1("B2"), chart.Series[0].ValuesTopLeft);
        Assert.Equal(CellPosition.FromA1("B4"), chart.Series[0].ValuesBottomRight);
        Assert.Equal(CellPosition.FromA1("C1"), chart.Series[1].Name);
    }

    [Fact]
    public void ChartBar()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Team"), Cell.Text("Score")),
                Row.Of(Cell.Text("Alpha"), Cell.Number(42)),
                Row.Of(Cell.Text("Beta"), Cell.Number(37)))
            .AddChart(ChartEntry.Of(ChartType.Bar, "A2", "A3", "D1", "K12", ChartSeries.Of("B1", "B2", "B3")));

        var loaded = VerifyScenario(nameof(ChartBar), Workbook.Create(sheet));
        var chart = Assert.Single(loaded.Sheets.Single().Charts);

        Assert.Equal(ChartType.Bar, chart.Type);
        Assert.Null(chart.Title);
        Assert.False(chart.ShowLegend);
        Assert.Single(chart.Series);
    }

    [Fact]
    public void Image()
    {
        var imageBytes = OnePixelGif();
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Logo below:")))
            .AddImage(ImageEntry.Of(imageBytes, ImageFormat.Gif, "A3", "C10"));

        var loaded = VerifyScenario(nameof(Image), Workbook.Create(sheet));
        var image = Assert.Single(loaded.Sheets.Single().Images);

        Assert.Equal(imageBytes, image.Data);
        Assert.Equal(ImageFormat.Gif, image.Format);
        Assert.Equal(CellPosition.FromA1("A3"), image.TopLeftAnchor);
        Assert.Equal(CellPosition.FromA1("C10"), image.BottomRightAnchor);
    }

    [Fact]
    public void PivotTableRowOnly()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Region"), Cell.Text("Sales")),
                Row.Of(Cell.Text("East"), Cell.Number(10)),
                Row.Of(Cell.Text("West"), Cell.Number(20)),
                Row.Of(Cell.Text("East"), Cell.Number(5)),
                Row.Of(Cell.Text("West"), Cell.Number(15)))
            .AddPivotTable(PivotTableEntry.Of("A1", "B5", "Region", "Sales", "D1"));

        var loaded = VerifyScenario(nameof(PivotTableRowOnly), Workbook.Create(sheet));
        var loadedSheet = loaded.Sheets.Single();
        var pivotTable = Assert.Single(loadedSheet.PivotTables);

        Assert.Null(pivotTable.SourceSheet);
        Assert.Equal(CellPosition.FromA1("A1"), pivotTable.SourceTopLeft);
        Assert.Equal(CellPosition.FromA1("B5"), pivotTable.SourceBottomRight);
        Assert.Equal("Region", pivotTable.RowField);
        Assert.Null(pivotTable.ColumnField);
        Assert.Equal("Sales", pivotTable.ValueField);
        Assert.Equal(PivotAggregation.Sum, pivotTable.Aggregation);
        Assert.Null(pivotTable.ValueCaption);
        Assert.Equal(CellPosition.FromA1("D1"), pivotTable.TopLeftAnchor);

        var numberAt = (string a1) =>
            Assert.IsType<CellValue.Number>(
                loadedSheet.Rows.SelectMany(r => r.Cells.Select(c => (Position: new CellPosition(r.Index!.Value, c.Column!.Value), c.Value)))
                    .Single(c => c.Position == CellPosition.FromA1(a1)).Value).Value;

        Assert.Equal(15.0, numberAt("E2"));
        Assert.Equal(35.0, numberAt("E3"));
        Assert.Equal(50.0, numberAt("E4"));
    }

    [Fact]
    public void PivotTableRowAndColumn()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Region"), Cell.Text("Quarter"), Cell.Text("Sales")),
                Row.Of(Cell.Text("East"), Cell.Text("Q1"), Cell.Number(10)),
                Row.Of(Cell.Text("East"), Cell.Text("Q2"), Cell.Number(5)),
                Row.Of(Cell.Text("West"), Cell.Text("Q1"), Cell.Number(20)),
                Row.Of(Cell.Text("West"), Cell.Text("Q2"), Cell.Number(15)))
            .AddPivotTable(
                PivotTableEntry
                    .Of("A1", "C5", "Region", "Sales", "E1")
                    .WithColumnField("Quarter")
                    .WithValueCaption("Total Sales"));

        var loaded = VerifyScenario(nameof(PivotTableRowAndColumn), Workbook.Create(sheet));
        var loadedSheet = loaded.Sheets.Single();
        var pivotTable = Assert.Single(loadedSheet.PivotTables);

        Assert.Equal("Quarter", pivotTable.ColumnField);
        Assert.Equal("Total Sales", pivotTable.ValueCaption);

        var numberAt = (string a1) =>
            Assert.IsType<CellValue.Number>(
                loadedSheet.Rows.SelectMany(r => r.Cells.Select(c => (Position: new CellPosition(r.Index!.Value, c.Column!.Value), c.Value)))
                    .Single(c => c.Position == CellPosition.FromA1(a1)).Value).Value;

        // E1 Region | F1 Q1 | G1 Q2 | H1 Grand Total
        // E2 East   | F2 10 | G2 5  | H2 15
        // E3 West   | F3 20 | G3 15 | H3 35
        // E4 Grand Total | F4 30 | G4 20 | H4 50
        Assert.Equal(10.0, numberAt("F2"));
        Assert.Equal(5.0, numberAt("G2"));
        Assert.Equal(15.0, numberAt("H2"));
        Assert.Equal(20.0, numberAt("F3"));
        Assert.Equal(15.0, numberAt("G3"));
        Assert.Equal(35.0, numberAt("H3"));
        Assert.Equal(30.0, numberAt("F4"));
        Assert.Equal(20.0, numberAt("G4"));
        Assert.Equal(50.0, numberAt("H4"));
    }

    [Fact]
    public void PivotTableAcrossSheets()
    {
        var sourceSheet = Sheet.Create(
            "Data",
            Row.Of(Cell.Text("Category"), Cell.Text("Amount")),
            Row.Of(Cell.Text("A"), Cell.Number(3)),
            Row.Of(Cell.Text("B"), Cell.Number(7)),
            Row.Of(Cell.Text("A"), Cell.Number(4)));

        var reportSheet = Sheet
            .Create("Report", Row.Of(Cell.Text("Pivot table below:")))
            .AddPivotTable(
                PivotTableEntry
                    .Of("A1", "B4", "Category", "Amount", "A3")
                    .WithSourceSheet("Data")
                    .WithAggregation(PivotAggregation.Count));

        var workbook = Workbook.Create(sourceSheet).AddSheet(reportSheet);

        var loaded = VerifyScenario(nameof(PivotTableAcrossSheets), workbook);
        var loadedReportSheet = loaded.Sheets.Single(s => s.Name == "Report");
        var pivotTable = Assert.Single(loadedReportSheet.PivotTables);

        Assert.Equal("Data", pivotTable.SourceSheet);
        Assert.Equal(PivotAggregation.Count, pivotTable.Aggregation);

        var numberAt = (string a1) =>
            Assert.IsType<CellValue.Number>(
                loadedReportSheet.Rows.SelectMany(r => r.Cells.Select(c => (Position: new CellPosition(r.Index!.Value, c.Column!.Value), c.Value)))
                    .Single(c => c.Position == CellPosition.FromA1(a1)).Value).Value;

        // A3 Category | B3 Count of Amount
        // A4 A        | B4 2
        // A5 B        | B5 1
        // A6 Grand Total | B6 3
        Assert.Equal(2.0, numberAt("B4"));
        Assert.Equal(1.0, numberAt("B5"));
        Assert.Equal(3.0, numberAt("B6"));
    }

    [Fact]
    public void VbaMacro()
    {
        var vbaProject = SampleVbaProject();
        var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Text("Run the HelloWorld macro")));
        var workbook = Workbook.Create(sheet).WithVbaProject(vbaProject);

        var loaded = VerifyScenario(nameof(VbaMacro), workbook, "output.xlsm");
        Assert.NotNull(loaded.VbaProject);
        Assert.Equal(vbaProject, loaded.VbaProject);
    }

    [Fact]
    public void SparklinesLineGroup()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Widgets"), Cell.Number(3), Cell.Number(8), Cell.Number(5), Cell.Number(9)),
                Row.Of(Cell.Text("Gadgets"), Cell.Number(6), Cell.Number(4), Cell.Number(7), Cell.Number(2)))
            .AddSparklineGroup(
                new SparklineGroupEntry(
                        SparklineCell.Of("F1", "B1", "E1"),
                        SparklineCell.Of("F2", "B2", "E2"))
                    .WithStyle(SparklineStyle.Default.WithHigh().WithLow()));

        var loaded = VerifyScenario(nameof(SparklinesLineGroup), Workbook.Create(sheet));
        var group = Assert.Single(loaded.Sheets.Single().SparklineGroups);

        Assert.Equal(SparklineType.Line, group.Style.Type);
        Assert.True(group.Style.ShowHigh);
        Assert.True(group.Style.ShowLow);
        Assert.Null(group.Style.Color);

        Assert.Equal(2, group.Sparklines.Count);
        Assert.Equal(CellPosition.FromA1("F1"), group.Sparklines[0].Cell);
        Assert.Equal(CellPosition.FromA1("B1"), group.Sparklines[0].DataTopLeft);
        Assert.Equal(CellPosition.FromA1("E1"), group.Sparklines[0].DataBottomRight);
        Assert.Equal(CellPosition.FromA1("F2"), group.Sparklines[1].Cell);
    }

    [Fact]
    public void SparklinesColumnGroup()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(-2), Cell.Number(4), Cell.Number(-1), Cell.Number(3)))
            .AddSparklineGroup(
                new SparklineGroupEntry(SparklineCell.Of("E1", "A1", "D1"))
                    .WithStyle(
                        SparklineStyle.Default
                            .WithType(SparklineType.Column)
                            .WithColor(new RgbColor(0, 112, 192))
                            .WithNegative()));

        var loaded = VerifyScenario(nameof(SparklinesColumnGroup), Workbook.Create(sheet));
        var group = Assert.Single(loaded.Sheets.Single().SparklineGroups);

        Assert.Equal(SparklineType.Column, group.Style.Type);
        Assert.Equal(new RgbColor(0, 112, 192), group.Style.Color);
        Assert.True(group.Style.ShowNegative);
        Assert.Single(group.Sparklines);
    }

    /// <summary>Matches the F# suite's own <c>redFillStyle</c>/<c>greenFillStyle</c> fixtures
    /// exactly, shared across the conditional formatting scenarios below.</summary>
    private static readonly CellStyle RedFillStyle = CellStyle.Default.WithFillColor(new RgbColor(255, 199, 206));

    private static readonly CellStyle GreenFillStyle = CellStyle.Default.WithFillColor(new RgbColor(198, 239, 206));

    [Fact]
    public void ConditionalFormat_CellValueRule()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(50)), Row.Of(Cell.Number(150)), Row.Of(Cell.Number(90)))
            .AddConditionalFormat(
                ConditionalFormatEntry.Of(
                    "A1",
                    "A3",
                    new ConditionalFormatRule.CellValueRule(ComparisonOperator.GreaterThan, "100", null, RedFillStyle)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_CellValueRule), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        Assert.Equal(CellPosition.FromA1("A1"), entry.TopLeft);
        Assert.Equal(CellPosition.FromA1("A3"), entry.BottomRight);
        var rule = Assert.IsType<ConditionalFormatRule.CellValueRule>(entry.Rule);
        Assert.Equal(ComparisonOperator.GreaterThan, rule.Operator);
        Assert.Equal("100", rule.Formula1);
        Assert.Null(rule.Formula2);
        Assert.Equal(new RgbColor(255, 199, 206), rule.Style.FillColor);
    }

    [Fact]
    public void ConditionalFormat_FormulaRule()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(10), Cell.Number(20)), Row.Of(Cell.Number(30), Cell.Number(5)))
            .AddConditionalFormat(
                ConditionalFormatEntry.Of("A1", "A2", new ConditionalFormatRule.FormulaRule("A1>B1", GreenFillStyle)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_FormulaRule), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.FormulaRule>(entry.Rule);
        Assert.Equal("A1>B1", rule.Formula);
        Assert.Equal(new RgbColor(198, 239, 206), rule.Style.FillColor);
    }

    [Fact]
    public void ConditionalFormat_ColorScale2()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(10)), Row.Of(Cell.Number(50)), Row.Of(Cell.Number(90)))
            .AddConditionalFormat(
                ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.ColorScale2(RgbColor.White, RgbColor.Red)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_ColorScale2), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.ColorScale2>(entry.Rule);
        Assert.Equal(RgbColor.White, rule.MinColor);
        Assert.Equal(RgbColor.Red, rule.MaxColor);
    }

    [Fact]
    public void ConditionalFormat_ColorScale3()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(10)), Row.Of(Cell.Number(50)), Row.Of(Cell.Number(90)))
            .AddConditionalFormat(
                ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.ColorScale3(RgbColor.Red, RgbColor.Yellow, RgbColor.Green)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_ColorScale3), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.ColorScale3>(entry.Rule);
        Assert.Equal(RgbColor.Red, rule.MinColor);
        Assert.Equal(RgbColor.Yellow, rule.MidColor);
        Assert.Equal(RgbColor.Green, rule.MaxColor);
    }

    [Fact]
    public void ConditionalFormat_DataBar()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Number(10)), Row.Of(Cell.Number(50)), Row.Of(Cell.Number(90)))
            .AddConditionalFormat(ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.DataBarRule(RgbColor.Blue)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_DataBar), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.DataBarRule>(entry.Rule);
        Assert.Equal(RgbColor.Blue, rule.Color);
    }

    [Fact]
    public void ConditionalFormat_DuplicateValues()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Apple")), Row.Of(Cell.Text("Banana")), Row.Of(Cell.Text("Apple")))
            .AddConditionalFormat(ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.DuplicateValuesRule(RedFillStyle)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_DuplicateValues), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.DuplicateValuesRule>(entry.Rule);
        Assert.Equal(new RgbColor(255, 199, 206), rule.Style.FillColor);
    }

    [Fact]
    public void ConditionalFormat_UniqueValues()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Apple")), Row.Of(Cell.Text("Banana")), Row.Of(Cell.Text("Apple")))
            .AddConditionalFormat(ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.UniqueValuesRule(GreenFillStyle)));

        var loaded = VerifyScenario(nameof(ConditionalFormat_UniqueValues), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().ConditionalFormats);

        var rule = Assert.IsType<ConditionalFormatRule.UniqueValuesRule>(entry.Rule);
        Assert.Equal(new RgbColor(198, 239, 206), rule.Style.FillColor);
    }

    [Fact]
    public void DataValidation_List()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Size")))
            .AddDataValidation(DataValidationEntry.Of("A2", "A2", new ValidationKind.ListValidation("Small", "Medium", "Large")));

        var loaded = VerifyScenario(nameof(DataValidation_List), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.ListValidation>(entry.Kind);
        Assert.Equal(["Small", "Medium", "Large"], kind.Items);
    }

    [Fact]
    public void DataValidation_ListFromRange()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Small"), Cell.Text("Medium"), Cell.Text("Large")),
                Row.Of(Cell.Text("Size")))
            .AddDataValidation(DataValidationEntry.Of("A2", "A2", new ValidationKind.ListFromRangeValidation(CellPosition.FromA1("A1"), CellPosition.FromA1("C1"))));

        var loaded = VerifyScenario(nameof(DataValidation_ListFromRange), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.ListFromRangeValidation>(entry.Kind);
        Assert.Equal(CellPosition.FromA1("A1"), kind.TopLeft);
        Assert.Equal(CellPosition.FromA1("C1"), kind.BottomRight);
    }

    [Fact]
    public void DataValidation_WholeNumber()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Quantity")))
            .AddDataValidation(
                DataValidationEntry
                    .Of("A2", "A2", new ValidationKind.WholeNumberValidation(ComparisonOperator.GreaterThan, "0", null))
                    .WithAlert(
                        ValidationAlert.Default
                            .WithErrorTitle("Invalid quantity")
                            .WithErrorMessage("Quantity must be a positive whole number.")));

        var loaded = VerifyScenario(nameof(DataValidation_WholeNumber), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.WholeNumberValidation>(entry.Kind);
        Assert.Equal(ComparisonOperator.GreaterThan, kind.Operator);
        Assert.Equal("0", kind.Formula1);
        Assert.Equal("Invalid quantity", entry.Alert.ErrorTitle);
        Assert.Equal("Quantity must be a positive whole number.", entry.Alert.ErrorMessage);
    }

    [Fact]
    public void DataValidation_Decimal()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Fraction (0-1)")))
            .AddDataValidation(DataValidationEntry.Of("A2", "A2", new ValidationKind.DecimalValidation(ComparisonOperator.Between, "0", "1")));

        var loaded = VerifyScenario(nameof(DataValidation_Decimal), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.DecimalValidation>(entry.Kind);
        Assert.Equal(ComparisonOperator.Between, kind.Operator);
        Assert.Equal("0", kind.Formula1);
        Assert.Equal("1", kind.Formula2);
    }

    [Fact]
    public void DataValidation_TextLength()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Short code (<= 10 chars)")))
            .AddDataValidation(DataValidationEntry.Of("A2", "A2", new ValidationKind.TextLengthValidation(ComparisonOperator.LessThanOrEqual, "10", null)));

        var loaded = VerifyScenario(nameof(DataValidation_TextLength), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.TextLengthValidation>(entry.Kind);
        Assert.Equal(ComparisonOperator.LessThanOrEqual, kind.Operator);
        Assert.Equal("10", kind.Formula1);
        Assert.Null(kind.Formula2);
    }

    [Fact]
    public void DataValidation_Custom()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Must be a number")))
            .AddDataValidation(
                DataValidationEntry
                    .Of("A2", "A2", new ValidationKind.CustomValidation("ISNUMBER(A2)"))
                    .WithAlert(
                        ValidationAlert.Default
                            .WithAllowBlank(false)
                            .WithInputTitle("Note")
                            .WithInputMessage("Enter a numeric value.")));

        var loaded = VerifyScenario(nameof(DataValidation_Custom), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().DataValidations);

        var kind = Assert.IsType<ValidationKind.CustomValidation>(entry.Kind);
        Assert.Equal("ISNUMBER(A2)", kind.Formula);
        Assert.False(entry.Alert.AllowBlank);
        Assert.Equal("Note", entry.Alert.InputTitle);
        Assert.Equal("Enter a numeric value.", entry.Alert.InputMessage);
    }

    [Fact]
    public void Hyperlink_External()
    {
        var sheet = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Open-XML-SDK on GitHub")))
            .AddHyperlink(
                HyperlinkEntry
                    .Of("A1", new HyperlinkTarget.ExternalHyperlink("https://github.com/dotnet/Open-XML-SDK"))
                    .WithTooltip("Open in browser")
                    .WithDisplay("dotnet/Open-XML-SDK"));

        var loaded = VerifyScenario(nameof(Hyperlink_External), Workbook.Create(sheet));
        var entry = Assert.Single(loaded.Sheets.Single().Hyperlinks);

        Assert.Equal(CellPosition.FromA1("A1"), entry.TopLeft);
        Assert.Equal(CellPosition.FromA1("A1"), entry.BottomRight);
        var target = Assert.IsType<HyperlinkTarget.ExternalHyperlink>(entry.Target);
        Assert.Equal("https://github.com/dotnet/Open-XML-SDK", target.Url);
        Assert.Equal("Open in browser", entry.Tooltip);
        Assert.Equal("dotnet/Open-XML-SDK", entry.Display);
    }

    [Fact]
    public void Hyperlink_Internal()
    {
        var sheet1 = Sheet
            .Create("Sheet1", Row.Of(Cell.Text("Go to Sheet2")))
            .AddHyperlink(HyperlinkEntry.Of("A1", new HyperlinkTarget.InternalHyperlink("Sheet2!A1")));

        var sheet2 = Sheet.Create("Sheet2", Row.Of(Cell.Text("You made it!")));

        var loaded = VerifyScenario(nameof(Hyperlink_Internal), Workbook.Create(sheet1).AddSheet(sheet2));
        var loadedSheet1 = loaded.Sheets.Single(s => s.Name == "Sheet1");
        var entry = Assert.Single(loadedSheet1.Hyperlinks);

        var target = Assert.IsType<HyperlinkTarget.InternalHyperlink>(entry.Target);
        Assert.Equal("Sheet2!A1", target.Location);
        Assert.Null(entry.Tooltip);
        Assert.Null(entry.Display);
    }

    [Fact]
    public void Comments()
    {
        var sheet = Sheet
            .Create(
                "Sheet1",
                Row.Of(Cell.Text("Revenue"), Cell.Number(1250)),
                Row.Of(Cell.Text("Costs"), Cell.Number(900)))
            .AddComment(CommentEntry.Of("B1", "Figure is provisional pending audit.", "Alex"))
            .AddComment(CommentEntry.Of("B2", "Includes one-off relocation costs.", "Alex"))
            .AddComment(CommentEntry.Of("A1", "Double check this label."));

        var loaded = VerifyScenario(nameof(Comments), Workbook.Create(sheet));
        var comments = loaded.Sheets.Single().Comments;

        Assert.Equal(3, comments.Count);
        Assert.Contains(comments, c => c.Cell == CellPosition.FromA1("B1") && c.Author == "Alex" && c.Text == "Figure is provisional pending audit.");
        Assert.Contains(comments, c => c.Cell == CellPosition.FromA1("B2") && c.Author == "Alex" && c.Text == "Includes one-off relocation costs.");
        Assert.Contains(comments, c => c.Cell == CellPosition.FromA1("A1") && c.Author == "" && c.Text == "Double check this label.");
    }

    // --- Generated-script verification (slow: actually runs `dotnet run`) ------------------
    //
    // Every scenario above writes its own Examples/<name>/script.cs as a side effect of
    // VerifyScenario. This is the only place that script actually gets *executed* rather
    // than just generated - each one runs its scenario's script via `dotnet run --file`
    // and checks the regenerated output file round-trips to equivalent facts as the
    // committed one. Running `dotnet run` from cold is slow (a restore/build per process),
    // so this is its own Category=Slow group rather than part of the default `dotnet test`
    // loop - run it explicitly with:
    //   dotnet test --filter "Category=Slow"
    // The default fast loop is:
    //   dotnet test --filter "Category!=Slow"

    /// <summary>One scenario name per <c>Examples/</c> folder that has a <c>script.cs</c>
    /// (i.e. every scenario above, once its fast test has run at least once) - discovered
    /// from disk rather than hand-listed, so a new scenario is automatically covered
    /// without touching this list.</summary>
    public static TheoryData<string> ScenarioNames()
    {
        var data = new TheoryData<string>();
        var dir = ExamplesDir();

        if (Directory.Exists(dir))
            foreach (var scenarioDir in Directory.GetDirectories(dir))
                if (File.Exists(Path.Combine(scenarioDir, "script.cs")))
                    data.Add(Path.GetFileName(scenarioDir));

        return data;
    }

    /// <summary>Compares the facts this wrapper models between two loaded workbooks - not
    /// full deep equality (this assembly's records use array-backed collections, which
    /// don't get structural equality for free from the C# compiler the way F#'s own lists
    /// do), just the same shape of check every individual scenario test above already makes
    /// by hand: same sheets, same cell values in the same positions, same counts of every
    /// sheet-level fact.</summary>
    private static void AssertWorkbooksMatch(Workbook before, Workbook after)
    {
        Assert.Equal(before.Sheets.Select(s => s.Name), after.Sheets.Select(s => s.Name));

        foreach (var (b, a) in before.Sheets.Zip(after.Sheets))
        {
            static CellValue[] OrderedValues(Sheet s) =>
                s.Rows.OrderBy(r => r.Index).SelectMany(r => r.Cells.OrderBy(c => c.Column).Select(c => c.Value)).ToArray();

            Assert.Equal(OrderedValues(b), OrderedValues(a));
            Assert.Equal(b.MergedRanges.Count, a.MergedRanges.Count);
            Assert.Equal(b.FreezePane is not null, a.FreezePane is not null);
            Assert.Equal(b.AutoFilter is not null, a.AutoFilter is not null);
            Assert.Equal(b.Tables.Count, a.Tables.Count);
            Assert.Equal(b.Charts.Count, a.Charts.Count);
            Assert.Equal(b.Images.Count, a.Images.Count);
            Assert.Equal(b.PivotTables.Count, a.PivotTables.Count);
        }

        Assert.Equal(before.VbaProject, after.VbaProject);
    }

    [Theory]
    [Trait("Category", "Slow")]
    [MemberData(nameof(ScenarioNames))]
    public void Example_script_regenerates_an_equivalent_file(string name)
    {
        var dir = Path.Combine(ExamplesDir(), name);
        var outputPath = Directory.GetFiles(dir).Single(f => !f.EndsWith("script.cs", StringComparison.Ordinal));

        var before = WorkbookIO.Load(outputPath);

        // The OpenXml SDK's read side can leave a memory-mapped view of the file lingering
        // past Dispose on Windows (a documented SDK quirk) - without forcing it closed
        // here, the `dotnet run` subprocess below can fail to reopen the same path.
        GC.Collect();
        GC.WaitForPendingFinalizers();

        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "dotnet",
            ArgumentList = { "run", "--file", "script.cs" },
            WorkingDirectory = dir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        })!;

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        Assert.True(process.ExitCode == 0, $"dotnet run script.cs ({name}) failed (exit {process.ExitCode}):\n{stdout}\n{stderr}");

        AssertSchemaValid(outputPath);
        var after = WorkbookIO.Load(outputPath);

        AssertWorkbooksMatch(before, after);
    }
}
