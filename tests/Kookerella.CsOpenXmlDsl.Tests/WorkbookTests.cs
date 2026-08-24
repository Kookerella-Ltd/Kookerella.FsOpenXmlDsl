using System.Diagnostics;
using Kookerella.CsOpenXmlDsl;
using Xunit;
using static Kookerella.CsOpenXmlDsl.Tests.TestHelpers;

namespace Kookerella.CsOpenXmlDsl.Tests;

public class WorkbookTests
{
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
    public void Sheet_without_charts_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.Charts);

        var withChart = plain.AddChart(ChartEntry.Of(ChartType.Line, "A2", "A3", "D1", "K12", ChartSeries.Of("B1", "B2", "B3")));
        Assert.Empty(plain.Charts);
        Assert.Single(withChart.Charts);
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
    public void Sheet_without_pivot_tables_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.PivotTables);

        var withPivotTable = plain.AddPivotTable(PivotTableEntry.Of("A1", "B2", "Region", "Sales", "D1"));
        Assert.Empty(plain.PivotTables);
        Assert.Single(withPivotTable.PivotTables);
    }

    [Fact]
    public void Sheet_without_sparkline_groups_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.SparklineGroups);

        var withGroup = plain.AddSparklineGroup(new SparklineGroupEntry(SparklineCell.Of("B1", "A1", "A1")));
        Assert.Empty(plain.SparklineGroups);
        Assert.Single(withGroup.SparklineGroups);
    }

    [Fact]
    public void SparklineStyle_fluent_methods_do_not_mutate_the_original_instance()
    {
        var original = SparklineStyle.Default;
        var withHigh = original.WithHigh();

        Assert.False(original.ShowHigh);
        Assert.True(withHigh.ShowHigh);
        Assert.NotSame(original, withHigh);
    }

    [Fact]
    public void Sheet_without_conditional_formats_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.ConditionalFormats);

        var withFormat = plain.AddConditionalFormat(
            ConditionalFormatEntry.Of("A1", "A1", new ConditionalFormatRule.DuplicateValuesRule(CellStyle.Default)));
        Assert.Empty(plain.ConditionalFormats);
        Assert.Single(withFormat.ConditionalFormats);
    }

    [Fact]
    public void Sheet_without_data_validations_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.DataValidations);

        var withValidation = plain.AddDataValidation(DataValidationEntry.Of("A1", "A1", new ValidationKind.CustomValidation("TRUE")));
        Assert.Empty(plain.DataValidations);
        Assert.Single(withValidation.DataValidations);
    }

    [Fact]
    public void ValidationAlert_fluent_methods_do_not_mutate_the_original_instance()
    {
        var original = ValidationAlert.Default;
        var withoutBlank = original.WithAllowBlank(false);

        Assert.True(original.AllowBlank);
        Assert.False(withoutBlank.AllowBlank);
        Assert.NotSame(original, withoutBlank);
    }

    [Fact]
    public void Sheet_without_hyperlinks_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.Hyperlinks);

        var withHyperlink = plain.AddHyperlink(HyperlinkEntry.Of("A1", new HyperlinkTarget.ExternalHyperlink("https://example.com")));
        Assert.Empty(plain.Hyperlinks);
        Assert.Single(withHyperlink.Hyperlinks);
    }

    [Fact]
    public void Sheet_without_comments_defaults_to_empty_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Empty(plain.Comments);

        var withComment = plain.AddComment(CommentEntry.Of("A1", "A note"));
        Assert.Empty(plain.Comments);
        Assert.Single(withComment.Comments);
    }

    [Fact]
    public void CommentEntry_author_defaults_to_empty_string()
    {
        var comment = CommentEntry.Of("A1", "A note");
        Assert.Equal("", comment.Author);
    }

    [Fact]
    public void Sheet_without_page_setup_defaults_to_null_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Null(plain.PageSetup);

        var withPageSetup = plain.WithPageSetup(PageSetup.Default.WithOrientation(PageOrientation.Landscape));
        Assert.Null(plain.PageSetup);
        Assert.Equal(PageOrientation.Landscape, withPageSetup.PageSetup!.Orientation);
    }

    [Fact]
    public void PageSetup_fluent_methods_do_not_mutate_the_original_instance()
    {
        var original = PageSetup.Default;
        var landscape = original.WithOrientation(PageOrientation.Landscape);

        Assert.Equal(PageOrientation.Portrait, original.Orientation);
        Assert.Equal(PageOrientation.Landscape, landscape.Orientation);
        Assert.NotSame(original, landscape);
    }

    [Fact]
    public void PageMargins_default_matches_Excels_own_built_in_margins()
    {
        Assert.Equal(0.7, PageMargins.Default.Left);
        Assert.Equal(0.7, PageMargins.Default.Right);
        Assert.Equal(0.75, PageMargins.Default.Top);
        Assert.Equal(0.75, PageMargins.Default.Bottom);
        Assert.Equal(0.3, PageMargins.Default.Header);
        Assert.Equal(0.3, PageMargins.Default.Footer);
    }

    [Fact]
    public void Workbook_without_defined_names_defaults_to_empty_and_is_immutable()
    {
        var plain = Workbook.Create();
        Assert.Empty(plain.DefinedNames);

        var withDefinedName = plain.AddDefinedName(DefinedNameEntry.Of("TaxRate", "0.075"));
        Assert.Empty(plain.DefinedNames);
        Assert.Single(withDefinedName.DefinedNames);
    }

    [Fact]
    public void DefinedNameEntry_hidden_defaults_to_false()
    {
        var definedName = DefinedNameEntry.Of("TaxRate", "0.075");
        Assert.False(definedName.Hidden);
    }

    [Fact]
    public void Sheet_without_protection_defaults_to_null_and_is_immutable()
    {
        var plain = Sheet.Create("Sheet1");
        Assert.Null(plain.Protection);

        var withProtection = plain.WithProtection(SheetProtection.Default.WithFormatCellsBlocked());
        Assert.Null(plain.Protection);
        Assert.True(withProtection.Protection!.FormatCellsBlocked);
    }

    [Fact]
    public void Workbook_without_protection_defaults_to_null_and_is_immutable()
    {
        var plain = Workbook.Create();
        Assert.Null(plain.Protection);

        var withProtection = plain.WithProtection(WorkbookProtection.Default.WithLockStructure());
        Assert.Null(plain.Protection);
        Assert.True(withProtection.Protection!.LockStructure);
    }

    [Fact]
    public void SheetProtection_fluent_methods_do_not_mutate_the_original_instance()
    {
        var original = SheetProtection.Default;
        var blocked = original.WithFormatCellsBlocked();

        Assert.Null(original.FormatCellsBlocked);
        Assert.True(blocked.FormatCellsBlocked);
        Assert.NotSame(original, blocked);
    }

    [Fact]
    public void SheetProtection_default_protects_the_sheet_with_no_other_flags_set()
    {
        Assert.True(SheetProtection.Default.Sheet);
        Assert.Null(SheetProtection.Default.Password);
        Assert.Null(SheetProtection.Default.FormatCellsBlocked);
        Assert.Null(SheetProtection.Default.SelectUnlockedCellsBlocked);
    }

    [Fact]
    public void Generate_produces_readable_source_for_a_simple_workbook()
    {
        var sheet = Sheet.Create(
            "Sheet1",
            Row.Of(Cell.Text("Item"), Cell.Number(42.5).WithStyle(CellStyle.Default.AsBold())));

        var script = CsCodeGen.Generate(["#:project ../Kookerella.CsOpenXmlDsl.csproj"], "out.xlsx", Workbook.Create(sheet));

        Assert.StartsWith("#:project ../Kookerella.CsOpenXmlDsl.csproj", script);
        Assert.Contains("using Kookerella.CsOpenXmlDsl;", script);
        Assert.Contains("Sheet.Create(", script);
        Assert.Contains("Cell.Text(\"Item\")", script);
        Assert.Contains("Cell.Number(42.5).WithStyle(CellStyle.Default.AsBold())", script);
        Assert.Contains("var workbook = Workbook.Create(sheet0);", script);
        Assert.Contains("WorkbookIO.Save(workbook, \"out.xlsx\");", script);

        // No cell in this sheet needs an explicit AtIndex/AtColumn - both rows/columns are
        // already sequential, so nothing should mention either.
        Assert.DoesNotContain("AtIndex", script);
        Assert.DoesNotContain("AtColumn", script);
    }

    [Fact]
    public void Generate_only_emits_AtIndex_and_AtColumn_where_positions_deviate_from_sequential()
    {
        var sheet = Sheet.Create(
            "Sheet1",
            Row.Of(Cell.Text("A1"), Cell.Text("C1").AtColumn(2)),
            Row.Of(Cell.Text("A5")).AtIndex(4));

        var script = CsCodeGen.Generate([], "out.xlsx", Workbook.Create(sheet));

        Assert.Contains("Cell.Text(\"A1\")", script);
        Assert.Contains("Cell.Text(\"C1\").AtColumn(2)", script);
        Assert.Contains("Row.Of(Cell.Text(\"A5\")).AtIndex(4)", script);
    }

    [Fact]
    public void Generate_omits_WithVbaProject_when_the_workbook_has_none()
    {
        var script = CsCodeGen.Generate([], "out.xlsx", Workbook.Create(Sheet.Create("Sheet1")));
        Assert.DoesNotContain("WithVbaProject", script);
    }

    [Fact]
    [Trait("Category", "Slow")]
    public void Generated_script_regenerates_an_equivalent_file()
    {
        var csprojPath = Path.Combine(FindRepoRoot(), "src", "Kookerella.CsOpenXmlDsl", "Kookerella.CsOpenXmlDsl.csproj");
        var scriptPath = Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslCodeGen_{Guid.NewGuid():N}.cs");
        var outputPath = TempXlsxPath();

        try
        {
            var headerStyle = CellStyle.Default.AsBold().WithFillColor(new RgbColor(220, 220, 220));

            var sheet = Sheet
                .Create(
                    "Sheet1",
                    Row.Of(Cell.Text("Item").WithStyle(headerStyle), Cell.Text("Qty").WithStyle(headerStyle)),
                    Row.Of(Cell.Text("Widgets"), Cell.Number(4), Cell.Formula("B2*2", 8.0)))
                .WithMergedRanges(MergedRange.Of("A1", "A1"))
                .WithFreezePane(1, 0)
                .WithAutoFilter(AutoFilterRange.Of("A1", "B2"))
                .WithTables(TableEntry.Of("A1", "B2", "Inventory", new TableColumn("Item"), new TableColumn("Qty")))
                .AddChart(
                    ChartEntry
                        .Of(ChartType.Column, "A2", "A2", "D1", "K12", ChartSeries.Of("B1", "B2", "B2"))
                        .WithTitle("Chart")
                        .WithLegend())
                .AddImage(ImageEntry.Of(OnePixelGif(), ImageFormat.Gif, "D15", "F20"))
                .AddPivotTable(PivotTableEntry.Of("A1", "B2", "Item", "Qty", "D25").WithAggregation(PivotAggregation.Sum));

            var workbook = Workbook.Create(sheet);

            var script = CsCodeGen.Generate([$"#:project {csprojPath}"], outputPath, workbook);
            File.WriteAllText(scriptPath, script);

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "dotnet",
                ArgumentList = { "run", "--file", scriptPath },
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            })!;

            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            Assert.True(process.ExitCode == 0, $"dotnet run {scriptPath} failed (exit {process.ExitCode}):\n{stdout}\n{stderr}");
            AssertSchemaValid(outputPath);

            var loaded = WorkbookIO.Load(outputPath);
            var loadedSheet = Assert.Single(loaded.Sheets);

            Assert.Equal("Sheet1", loadedSheet.Name);
            Assert.Single(loadedSheet.MergedRanges);
            Assert.NotNull(loadedSheet.FreezePane);
            Assert.NotNull(loadedSheet.AutoFilter);
            Assert.Single(loadedSheet.Tables);
            Assert.Single(loadedSheet.Charts);
            Assert.Single(loadedSheet.Images);
            Assert.Single(loadedSheet.PivotTables);

            var loadedTable = loadedSheet.Tables.Single();
            Assert.Equal("Inventory", loadedTable.Name);

            var loadedChart = loadedSheet.Charts.Single();
            Assert.Equal("Chart", loadedChart.Title);
            Assert.True(loadedChart.ShowLegend);

            Assert.Equal(OnePixelGif(), loadedSheet.Images.Single().Data);

            var loadedPivotTable = loadedSheet.PivotTables.Single();
            Assert.Equal(PivotAggregation.Sum, loadedPivotTable.Aggregation);
        }
        finally
        {
            if (File.Exists(scriptPath))
                File.Delete(scriptPath);
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }
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
