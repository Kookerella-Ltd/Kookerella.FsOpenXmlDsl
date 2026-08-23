using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Kookerella.CsOpenXmlDsl;
using Xunit;

namespace Kookerella.CsOpenXmlDsl.Tests;

public class WorkbookTests
{
    private static string TempXlsxPath() =>
        Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslTest_{Guid.NewGuid():N}.xlsx");

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
}
