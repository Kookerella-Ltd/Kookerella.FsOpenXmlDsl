using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Kookerella.CsOpenXmlDsl;
using Xunit;

namespace Kookerella.CsOpenXmlDsl.Tests;

public class WorkbookTests
{
    private static string TempXlsxPath() =>
        Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslTest_{Guid.NewGuid():N}.xlsx");

    private static string TempXlsmPath() =>
        Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslTest_{Guid.NewGuid():N}.xlsm");

    /// <summary>A real vbaProject.bin, shared with the F# test suite (extracted from a
    /// workbook actually saved by Excel) - see this project's own .csproj for how it's
    /// linked in rather than copy-pasted.</summary>
    private static byte[] SampleVbaProject() =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", "sample.vbaProject.bin"));

    /// <summary>The canonical "1x1 transparent GIF" - the smallest possible valid image
    /// file, used ubiquitously as a web tracking pixel, so its bytes are about as
    /// well-known and trustworthy as test fixtures get. Same fixture the F# suite uses.</summary>
    private static byte[] OnePixelGif() =>
        Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7");

    private static void AssertSchemaValid(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void Cell_values_and_formulas_round_trip()
    {
        var path = TempXlsxPath();
        try
        {
            var sheet = Sheet.Create(
                "Sheet1",
                Row.Of(Cell.Text("Item"), Cell.Text("Qty"), Cell.Text("Price"), Cell.Text("Total")),
                Row.Of(Cell.Text("Widgets"), Cell.Number(4), Cell.Number(2.5), Cell.Formula("B2*C2", 10.0)),
                Row.Of(Cell.Boolean(true), Cell.Date(new DateTime(2026, 1, 1))));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
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
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Explicit_row_and_column_positions_are_respected()
    {
        var path = TempXlsxPath();
        try
        {
            // Row 0 implicit, row 4 explicit (gap at 1-3), column 0 implicit then column 2 explicit.
            var sheet = Sheet.Create(
                "Sheet1",
                Row.Of(Cell.Text("A1"), Cell.Text("B1")),
                Row.Of(Cell.Text("A5"), Cell.Text("C5").AtColumn(2)).AtIndex(4));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            var loaded = WorkbookIO.Load(path);
            var loadedSheet = Assert.Single(loaded.Sheets);

            var allCells = loadedSheet.Rows.SelectMany(r => r.Cells.Select(c => (Row: r.Index, Col: c.Column, c.Value)));
            Assert.Contains(allCells, c => c.Row == 0 && c.Col == 0 && ((CellValue.Text)c.Value!).Value == "A1");
            Assert.Contains(allCells, c => c.Row == 0 && c.Col == 1 && ((CellValue.Text)c.Value!).Value == "B1");
            Assert.Contains(allCells, c => c.Row == 4 && c.Col == 0 && ((CellValue.Text)c.Value!).Value == "A5");
            Assert.Contains(allCells, c => c.Row == 4 && c.Col == 2 && ((CellValue.Text)c.Value!).Value == "C5");
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Styling_round_trips()
    {
        var path = TempXlsxPath();
        try
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

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
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
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Custom_number_format_round_trips()
    {
        var path = TempXlsxPath();
        try
        {
            var style = CellStyle.Default.WithCustomNumberFormat("0.000%");
            var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Number(0.5).WithStyle(style)));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            var loaded = WorkbookIO.Load(path);
            var cell = loaded.Sheets.Single().Rows.Single().Cells.Single();

            Assert.Null(cell.Style!.NumberFormat);
            Assert.Equal("0.000%", cell.Style.CustomNumberFormat);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Multiple_sheets_round_trip()
    {
        var path = TempXlsxPath();
        try
        {
            var wb = Workbook
                .Create(Sheet.Create("First", Row.Of(Cell.Text("a"))))
                .AddSheet(Sheet.Create("Second", Row.Of(Cell.Text("b"))));

            WorkbookIO.Save(wb, path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            Assert.Equal(["First", "Second"], loaded.Sheets.Select(s => s.Name));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Fluent_methods_do_not_mutate_the_original_instance()
    {
        var original = CellStyle.Default;
        var bolded = original.AsBold();

        Assert.False(original.Bold);
        Assert.True(bolded.Bold);
        Assert.NotSame(original, bolded);

        var baseCell = Cell.Text("x");
        var styled = baseCell.WithStyle(bolded);
        Assert.Null(baseCell.Style);
        Assert.Same(bolded, styled.Style);

        var row = Row.Of(Cell.Text("a"));
        var reindexed = row.AtIndex(5);
        Assert.Null(row.Index);
        Assert.Equal(5, reindexed.Index);

        var sheet = Sheet.Create("S");
        var withRow = sheet.AddRow(row);
        Assert.Empty(sheet.Rows);
        Assert.Single(withRow.Rows);

        var workbook = Workbook.Create();
        var withSheet = workbook.AddSheet(sheet);
        Assert.Empty(workbook.Sheets);
        Assert.Single(withSheet.Sheets);
    }

    [Fact]
    public void Save_and_load_via_streams_round_trip()
    {
        using var stream = new MemoryStream();
        var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Text("hello")));

        WorkbookIO.Save(Workbook.Create(sheet), stream);
        stream.Position = 0;

        var loaded = WorkbookIO.Load(stream);
        var cell = loaded.Sheets.Single().Rows.Single().Cells.Single();
        Assert.Equal("hello", ((CellValue.Text)cell.Value).Value);
    }

    [Fact]
    public void Merged_ranges_freeze_pane_and_autofilter_round_trip()
    {
        var path = TempXlsxPath();
        try
        {
            var sheet = Sheet
                .Create(
                    "Sheet1",
                    Row.Of(Cell.Text("Quarterly Report")),
                    Row.Of(Cell.Text("Region"), Cell.Text("Sales")),
                    Row.Of(Cell.Text("East"), Cell.Number(10)))
                .WithMergedRanges(MergedRange.Of("A1", "B1"))
                .WithFreezePane(2, 0)
                .WithAutoFilter(AutoFilterRange.Of("A2", "B3"));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            var loadedSheet = loaded.Sheets.Single();

            var mergedRange = Assert.Single(loadedSheet.MergedRanges);
            Assert.Equal(new CellPosition(0, 0), mergedRange.TopLeft);
            Assert.Equal(new CellPosition(0, 1), mergedRange.BottomRight);
            Assert.Equal("A1", mergedRange.TopLeft.ToA1());
            Assert.Equal("B1", mergedRange.BottomRight.ToA1());

            Assert.NotNull(loadedSheet.FreezePane);
            Assert.Equal(2, loadedSheet.FreezePane!.Rows);
            Assert.Equal(0, loadedSheet.FreezePane.Columns);

            Assert.NotNull(loadedSheet.AutoFilter);
            Assert.Equal(CellPosition.FromA1("A2"), loadedSheet.AutoFilter!.TopLeft);
            Assert.Equal(CellPosition.FromA1("B3"), loadedSheet.AutoFilter.BottomRight);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Vba_project_round_trips_byte_for_byte()
    {
        var path = TempXlsmPath();
        try
        {
            var vbaProject = SampleVbaProject();
            var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Text("Run the HelloWorld macro")));
            var workbook = Workbook.Create(sheet).WithVbaProject(vbaProject);

            WorkbookIO.Save(workbook, path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            Assert.NotNull(loaded.VbaProject);
            Assert.Equal(vbaProject, loaded.VbaProject);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void WithVbaProject_defensively_copies_the_input_array()
    {
        var original = new byte[] { 1, 2, 3 };
        var workbook = Workbook.Create().WithVbaProject(original);

        original[0] = 99;

        Assert.Equal(new byte[] { 1, 2, 3 }, workbook.VbaProject);
    }

    [Fact]
    public void Workbook_without_a_vba_project_has_none()
    {
        var workbook = Workbook.Create(Sheet.Create("Sheet1"));
        Assert.Null(workbook.VbaProject);
    }

    [Fact]
    public void Sheet_level_facts_default_to_absent_and_are_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.MergedRanges);
        Assert.Null(plain.FreezePane);
        Assert.Null(plain.AutoFilter);

        var withMerge = plain.AddMergedRange(MergedRange.Of("A1", "A2"));
        Assert.Empty(plain.MergedRanges);
        Assert.Single(withMerge.MergedRanges);
    }

    [Fact]
    public void Table_round_trips()
    {
        var path = TempXlsxPath();
        try
        {
            var sheet = Sheet
                .Create(
                    "Sheet1",
                    Row.Of(Cell.Text("Item"), Cell.Text("Quantity")),
                    Row.Of(Cell.Text("Widgets"), Cell.Number(12)),
                    Row.Of(Cell.Text("Gadgets"), Cell.Number(5)))
                .WithTables(TableEntry.Of("A1", "B3", "Inventory", new TableColumn("Item"), new TableColumn("Quantity")));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            var table = Assert.Single(loaded.Sheets.Single().Tables);

            Assert.Equal("Inventory", table.Name);
            Assert.Equal(CellPosition.FromA1("A1"), table.TopLeft);
            Assert.Equal(CellPosition.FromA1("B3"), table.BottomRight);
            Assert.Equal(["Item", "Quantity"], table.Columns.Select(c => c.Name));
            Assert.All(table.Columns, c => Assert.Null(c.CalculatedFormula));
            Assert.Equal("TableStyleMedium2", table.Style.Name);
            Assert.True(table.Style.ShowRowStripes);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Table_with_calculated_column_and_custom_style_round_trips()
    {
        var path = TempXlsxPath();
        try
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

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            var table = Assert.Single(loaded.Sheets.Single().Tables);

            var total = table.Columns.Single(c => c.Name == "Total");
            Assert.Equal("[@Quantity]*[@[Unit Price]]", total.CalculatedFormula);
            Assert.Equal("TableStyleLight9", table.Style.Name);
            Assert.True(table.Style.ShowColumnStripes);
            Assert.False(table.Style.ShowRowStripes);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Chart_with_title_and_legend_round_trips()
    {
        var path = TempXlsxPath();
        try
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

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
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
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Chart_without_title_or_legend_round_trips()
    {
        var path = TempXlsxPath();
        try
        {
            var sheet = Sheet
                .Create(
                    "Sheet1",
                    Row.Of(Cell.Text("Team"), Cell.Text("Score")),
                    Row.Of(Cell.Text("Alpha"), Cell.Number(42)),
                    Row.Of(Cell.Text("Beta"), Cell.Number(37)))
                .AddChart(ChartEntry.Of(ChartType.Bar, "A2", "A3", "D1", "K12", ChartSeries.Of("B1", "B2", "B3")));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            var chart = Assert.Single(loaded.Sheets.Single().Charts);

            Assert.Equal(ChartType.Bar, chart.Type);
            Assert.Null(chart.Title);
            Assert.False(chart.ShowLegend);
            Assert.Single(chart.Series);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Sheet_without_charts_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.Charts);

        var withChart = plain.AddChart(ChartEntry.Of(ChartType.Line, "A2", "A3", "D1", "K12", ChartSeries.Of("B1", "B2", "B3")));
        Assert.Empty(plain.Charts);
        Assert.Single(withChart.Charts);
    }

    [Fact]
    public void Image_anchored_over_a_range_round_trips_byte_for_byte()
    {
        var path = TempXlsxPath();
        try
        {
            var imageBytes = OnePixelGif();
            var sheet = Sheet
                .Create("Sheet1", Row.Of(Cell.Text("Logo below:")))
                .AddImage(ImageEntry.Of(imageBytes, ImageFormat.Gif, "A3", "C10"));

            WorkbookIO.Save(Workbook.Create(sheet), path);
            AssertSchemaValid(path);

            var loaded = WorkbookIO.Load(path);
            var image = Assert.Single(loaded.Sheets.Single().Images);

            Assert.Equal(imageBytes, image.Data);
            Assert.Equal(ImageFormat.Gif, image.Format);
            Assert.Equal(CellPosition.FromA1("A3"), image.TopLeftAnchor);
            Assert.Equal(CellPosition.FromA1("C10"), image.BottomRightAnchor);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void ImageEntry_defensively_copies_the_input_array()
    {
        var original = new byte[] { 1, 2, 3 };
        var image = new ImageEntry(original, ImageFormat.Png, CellPosition.FromA1("A1"), CellPosition.FromA1("B2"));

        original[0] = 99;

        Assert.Equal(new byte[] { 1, 2, 3 }, image.Data);
    }

    [Fact]
    public void Sheet_without_images_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.Images);

        var withImage = plain.AddImage(ImageEntry.Of(OnePixelGif(), ImageFormat.Gif, "A1", "B2"));
        Assert.Empty(plain.Images);
        Assert.Single(withImage.Images);
    }

    [Fact]
    public void Table_with_mismatched_column_count_throws_on_save()
    {
        var path = TempXlsxPath();
        try
        {
            // Range is 2 columns wide (A1:B2) but only 1 TableColumn is given.
            var sheet = Sheet
                .Create("Sheet1", Row.Of(Cell.Text("A"), Cell.Text("B")), Row.Of(Cell.Text("1"), Cell.Text("2")))
                .WithTables(TableEntry.Of("A1", "B2", "Bad", new TableColumn("A")));

            Assert.Throws<ArgumentException>(() => WorkbookIO.Save(Workbook.Create(sheet), path));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
