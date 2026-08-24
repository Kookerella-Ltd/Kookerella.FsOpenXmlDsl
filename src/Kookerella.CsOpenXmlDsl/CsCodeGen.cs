using System.Globalization;
using System.Text;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Renders a <see cref="Workbook"/> back out as a self-contained C# file that regenerates
/// an equivalent file when run - the reverse of <see cref="WorkbookIO.Load(string)"/> one
/// level further: loading turns a file into this wrapper's types, this turns those types
/// into C# *source text*. Every renderer below is a direct, mechanical mirror of a type's
/// own fluent API (diffing against <c>.Default</c>/<see langword="null"/> where one exists,
/// so generated code only mentions what isn't already implied) - there's no separate
/// "codegen model", just string-building over this assembly's own public types.
/// <para>
/// Unlike the F# core's own <c>Workbook.generateScript</c> (which keeps every statement on
/// one line specifically to sidestep F#'s indentation-sensitive offside rule), C# has no
/// such constraint, so this emits normally-indented, human-readable source.
/// </para>
/// The emitted file targets .NET's "file-based apps" feature (<c>dotnet run script.cs</c>,
/// available from .NET 10) rather than a traditional project - the <c>referenceLines</c>
/// parameter on <see cref="Generate"/> is whatever raw <c>#:package</c>/<c>#:project</c>
/// directives the caller needs so the file can locate this assembly; this class has no
/// opinion on that, since it depends entirely on where the file ends up living relative to
/// the caller's own build (same reasoning the F# core's own doc comment on
/// <c>generateScript</c> gives for its analogous <c>referenceLines</c> parameter).
/// </summary>
public static class CsCodeGen
{
    private static string RenderString(string s) =>
        "\"" + s
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t")
        + "\"";

    /// <summary>Always includes a decimal point (or exponent) so the literal is
    /// unambiguously a C# <see langword="double"/>, not an <see langword="int"/>.</summary>
    private static string RenderDouble(double d)
    {
        if (double.IsNaN(d))
            return "double.NaN";
        if (double.IsPositiveInfinity(d))
            return "double.PositiveInfinity";
        if (double.IsNegativeInfinity(d))
            return "double.NegativeInfinity";

        var s = d.ToString(CultureInfo.InvariantCulture);
        return s.Contains('.') || s.Contains('E') ? s : s + ".0";
    }

    private static string RenderBool(bool b) => b ? "true" : "false";

    /// <summary>Renders from <see cref="DateTime.Ticks"/> rather than round-tripping
    /// through Excel's OADate convention (which the F# core's own codegen uses) - simpler,
    /// and exact, since this renders directly from an in-memory <see cref="DateTime"/>
    /// rather than from a value freshly read off disk.</summary>
    private static string RenderDateTime(DateTime d) => $"new DateTime({d.Ticks}L, DateTimeKind.{d.Kind})";

    private static string RenderCellPosition(CellPosition p) => $"CellPosition.FromA1({RenderString(p.ToA1())})";

    private static string RenderRgbColor(RgbColor c) => c switch
    {
        _ when c == RgbColor.Black => "RgbColor.Black",
        _ when c == RgbColor.White => "RgbColor.White",
        _ when c == RgbColor.Red => "RgbColor.Red",
        _ when c == RgbColor.Green => "RgbColor.Green",
        _ when c == RgbColor.Blue => "RgbColor.Blue",
        _ when c == RgbColor.Yellow => "RgbColor.Yellow",
        _ => $"new RgbColor({c.R}, {c.G}, {c.B})"
    };

    private static string RenderBorderSide(BorderSide side) =>
        side.Color is { } color
            ? $"new BorderSide(BorderLineStyle.{side.Style}, {RenderRgbColor(color)})"
            : $"new BorderSide(BorderLineStyle.{side.Style})";

    private static string? RenderCellBorder(CellBorder? border)
    {
        if (border is not { } b || b == CellBorder.None)
            return null;

        if (b.Left is { } left && left == b.Right && left == b.Top && left == b.Bottom)
            return $"CellBorder.None.WithAllSides({RenderBorderSide(left)})";

        var sb = new StringBuilder("CellBorder.None");
        if (b.Left is { } l)
            sb.Append($".WithLeft({RenderBorderSide(l)})");
        if (b.Right is { } r)
            sb.Append($".WithRight({RenderBorderSide(r)})");
        if (b.Top is { } t)
            sb.Append($".WithTop({RenderBorderSide(t)})");
        if (b.Bottom is { } bo)
            sb.Append($".WithBottom({RenderBorderSide(bo)})");

        return sb.ToString();
    }

    private static string RenderCellStyle(CellStyle style)
    {
        if (style == CellStyle.Default)
            return "CellStyle.Default";

        var sb = new StringBuilder("CellStyle.Default");

        if (style.FontName is { } fontName)
            sb.Append($".WithFontName({RenderString(fontName)})");
        if (style.FontSize is { } fontSize)
            sb.Append($".WithFontSize({RenderDouble(fontSize)})");
        if (style.Bold)
            sb.Append(".AsBold()");
        if (style.Italic)
            sb.Append(".AsItalic()");
        if (style.Underline)
            sb.Append(".AsUnderline()");
        if (style.Strikethrough)
            sb.Append(".AsStrikethrough()");
        if (style.FontColor is { } fontColor)
            sb.Append($".WithFontColor({RenderRgbColor(fontColor)})");
        if (style.FillColor is { } fillColor)
            sb.Append($".WithFillColor({RenderRgbColor(fillColor)})");
        if (RenderCellBorder(style.Border) is { } border)
            sb.Append($".WithBorder({border})");
        if (style.HorizontalAlignment is { } horizontal)
            sb.Append($".WithHorizontalAlignment(HorizontalCellAlignment.{horizontal})");
        if (style.VerticalAlignment is { } vertical)
            sb.Append($".WithVerticalAlignment(VerticalCellAlignment.{vertical})");
        if (style.WrapText)
            sb.Append(".AsWrapText()");
        if (style.NumberFormat is { } numberFormat)
            sb.Append($".WithNumberFormat(NumberFormatKind.{numberFormat})");
        if (style.CustomNumberFormat is { } customNumberFormat)
            sb.Append($".WithCustomNumberFormat({RenderString(customNumberFormat)})");

        return sb.ToString();
    }

    private static string RenderCellValueExpr(CellValue value) => value switch
    {
        CellValue.Text t => $"Cell.Text({RenderString(t.Value)})",
        CellValue.Number n => $"Cell.Number({RenderDouble(n.Value)})",
        CellValue.Boolean b => $"Cell.Boolean({RenderBool(b.Value)})",
        CellValue.Date d => $"Cell.Date({RenderDateTime(d.Value)})",
        CellValue.Formula f when f.CachedValue is { } cached => $"Cell.Formula({RenderString(f.Expression)}, {RenderDouble(cached)})",
        CellValue.Formula f => $"Cell.Formula({RenderString(f.Expression)})",
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, null)
    };

    private static string RenderCell(Cell cell, int? explicitColumn)
    {
        var expr = RenderCellValueExpr(cell.Value);

        if (explicitColumn is { } col)
            expr += $".AtColumn({col})";
        if (cell.Style is { } style && style != CellStyle.Default)
            expr += $".WithStyle({RenderCellStyle(style)})";

        return expr;
    }

    private static string RenderMergedRange(MergedRange range) =>
        $"MergedRange.Of({RenderString(range.TopLeft.ToA1())}, {RenderString(range.BottomRight.ToA1())})";

    private static string RenderFreezePane(FreezePane pane) => $"new FreezePane({pane.Rows}, {pane.Columns})";

    private static string RenderAutoFilterRange(AutoFilterRange range) =>
        $"AutoFilterRange.Of({RenderString(range.TopLeft.ToA1())}, {RenderString(range.BottomRight.ToA1())})";

    private static string RenderTableColumn(TableColumn column) =>
        column.CalculatedFormula is { } formula
            ? $"new TableColumn({RenderString(column.Name)}, {RenderString(formula)})"
            : $"new TableColumn({RenderString(column.Name)})";

    private static string RenderTableStyle(TableStyle style)
    {
        if (style == TableStyle.Default)
            return "TableStyle.Default";

        var sb = new StringBuilder("TableStyle.Default");

        if (style.Name != TableStyle.Default.Name)
            sb.Append($".WithName({(style.Name is { } name ? RenderString(name) : "null")})");
        if (style.ShowFirstColumn)
            sb.Append(".WithFirstColumnEmphasis()");
        if (style.ShowLastColumn)
            sb.Append(".WithLastColumnEmphasis()");
        if (!style.ShowRowStripes)
            sb.Append(".WithoutRowStripes()");
        if (style.ShowColumnStripes)
            sb.Append(".WithColumnStripes()");

        return sb.ToString();
    }

    private static string RenderTableEntry(TableEntry table)
    {
        var columns = string.Join(", ", table.Columns.Select(RenderTableColumn));
        var expr = $"TableEntry.Of({RenderString(table.TopLeft.ToA1())}, {RenderString(table.BottomRight.ToA1())}, {RenderString(table.Name)}, {columns})";

        return table.Style == TableStyle.Default ? expr : $"{expr}.WithStyle({RenderTableStyle(table.Style)})";
    }

    private static string RenderChartSeries(ChartSeries series) =>
        $"ChartSeries.Of({RenderString(series.Name.ToA1())}, {RenderString(series.ValuesTopLeft.ToA1())}, {RenderString(series.ValuesBottomRight.ToA1())})";

    private static string RenderChartEntry(ChartEntry chart)
    {
        var series = string.Join(", ", chart.Series.Select(RenderChartSeries));
        var sb = new StringBuilder(
            $"ChartEntry.Of(ChartType.{chart.Type}, {RenderString(chart.CategoriesTopLeft.ToA1())}, {RenderString(chart.CategoriesBottomRight.ToA1())}, {RenderString(chart.TopLeftAnchor.ToA1())}, {RenderString(chart.BottomRightAnchor.ToA1())}, {series})");

        if (chart.Title is { } title)
            sb.Append($".WithTitle({RenderString(title)})");
        if (chart.ShowLegend)
            sb.Append(".WithLegend()");

        return sb.ToString();
    }

    private static string RenderImageEntry(ImageEntry image) =>
        $"ImageEntry.Of(System.Convert.FromBase64String({RenderString(Convert.ToBase64String(image.Data))}), ImageFormat.{image.Format}, {RenderString(image.TopLeftAnchor.ToA1())}, {RenderString(image.BottomRightAnchor.ToA1())})";

    private static string RenderPivotTableEntry(PivotTableEntry pivotTable)
    {
        var sb = new StringBuilder(
            $"PivotTableEntry.Of({RenderString(pivotTable.SourceTopLeft.ToA1())}, {RenderString(pivotTable.SourceBottomRight.ToA1())}, {RenderString(pivotTable.RowField)}, {RenderString(pivotTable.ValueField)}, {RenderString(pivotTable.TopLeftAnchor.ToA1())})");

        if (pivotTable.SourceSheet is { } sourceSheet)
            sb.Append($".WithSourceSheet({RenderString(sourceSheet)})");
        if (pivotTable.ColumnField is { } columnField)
            sb.Append($".WithColumnField({RenderString(columnField)})");
        if (pivotTable.Aggregation != PivotAggregation.Sum)
            sb.Append($".WithAggregation(PivotAggregation.{pivotTable.Aggregation})");
        if (pivotTable.ValueCaption is { } valueCaption)
            sb.Append($".WithValueCaption({RenderString(valueCaption)})");

        return sb.ToString();
    }

    private static string RenderSparklineCell(SparklineCell cell) =>
        $"SparklineCell.Of({RenderString(cell.Cell.ToA1())}, {RenderString(cell.DataTopLeft.ToA1())}, {RenderString(cell.DataBottomRight.ToA1())})";

    private static string RenderSparklineStyle(SparklineStyle style)
    {
        if (style == SparklineStyle.Default)
            return "SparklineStyle.Default";

        var sb = new StringBuilder("SparklineStyle.Default");

        if (style.Type != SparklineStyle.Default.Type)
            sb.Append($".WithType(SparklineType.{style.Type})");
        if (style.Color is { } color)
            sb.Append($".WithColor({RenderRgbColor(color)})");
        if (style.LineWeight is { } lineWeight)
            sb.Append($".WithLineWeight({RenderDouble(lineWeight)})");
        if (style.ShowMarkers)
            sb.Append(".WithMarkers()");
        if (style.ShowHigh)
            sb.Append(".WithHigh()");
        if (style.ShowLow)
            sb.Append(".WithLow()");
        if (style.ShowFirst)
            sb.Append(".WithFirst()");
        if (style.ShowLast)
            sb.Append(".WithLast()");
        if (style.ShowNegative)
            sb.Append(".WithNegative()");

        return sb.ToString();
    }

    private static string RenderSparklineGroupEntry(SparklineGroupEntry group)
    {
        var sparklines = string.Join(", ", group.Sparklines.Select(RenderSparklineCell));
        var expr = $"new SparklineGroupEntry({sparklines})";

        return group.Style == SparklineStyle.Default ? expr : $"{expr}.WithStyle({RenderSparklineStyle(group.Style)})";
    }

    private static string RenderConditionalFormatRule(ConditionalFormatRule rule) => rule switch
    {
        ConditionalFormatRule.CellValueRule r =>
            $"new ConditionalFormatRule.CellValueRule(ComparisonOperator.{r.Operator}, {RenderString(r.Formula1)}, {(r.Formula2 is { } f2 ? RenderString(f2) : "null")}, {RenderCellStyle(r.Style)})",
        ConditionalFormatRule.FormulaRule r => $"new ConditionalFormatRule.FormulaRule({RenderString(r.Formula)}, {RenderCellStyle(r.Style)})",
        ConditionalFormatRule.ColorScale2 r => $"new ConditionalFormatRule.ColorScale2({RenderRgbColor(r.MinColor)}, {RenderRgbColor(r.MaxColor)})",
        ConditionalFormatRule.ColorScale3 r =>
            $"new ConditionalFormatRule.ColorScale3({RenderRgbColor(r.MinColor)}, {RenderRgbColor(r.MidColor)}, {RenderRgbColor(r.MaxColor)})",
        ConditionalFormatRule.DataBarRule r => $"new ConditionalFormatRule.DataBarRule({RenderRgbColor(r.Color)})",
        ConditionalFormatRule.DuplicateValuesRule r => $"new ConditionalFormatRule.DuplicateValuesRule({RenderCellStyle(r.Style)})",
        ConditionalFormatRule.UniqueValuesRule r => $"new ConditionalFormatRule.UniqueValuesRule({RenderCellStyle(r.Style)})",
        _ => throw new ArgumentOutOfRangeException(nameof(rule), rule, null)
    };

    private static string RenderConditionalFormatEntry(ConditionalFormatEntry entry) =>
        $"ConditionalFormatEntry.Of({RenderString(entry.TopLeft.ToA1())}, {RenderString(entry.BottomRight.ToA1())}, {RenderConditionalFormatRule(entry.Rule)})";

    /// <summary>Resolves the same <c>Index ?? nextRow</c>/<c>Column ?? nextColumn</c>
    /// convention <see cref="WorkbookIO"/> itself uses at save time, then only emits an
    /// explicit <c>.AtIndex</c>/<c>.AtColumn</c> where the resolved position actually
    /// deviates from strict sequential numbering - so generated code reads the way a human
    /// would write it by hand, the same principle the F# core's own codegen follows.</summary>
    private static string RenderRow(Row row, ref int nextRow)
    {
        var resolvedRowIndex = row.Index ?? nextRow;
        var explicitRowIndex = resolvedRowIndex == nextRow ? (int?)null : resolvedRowIndex;
        nextRow = resolvedRowIndex + 1;

        var nextColumn = 0;
        var cellRenders = new List<string>();

        foreach (var cell in row.Cells)
        {
            var resolvedColumn = cell.Column ?? nextColumn;
            var explicitColumn = resolvedColumn == nextColumn ? (int?)null : resolvedColumn;
            cellRenders.Add(RenderCell(cell, explicitColumn));
            nextColumn = resolvedColumn + 1;
        }

        var rowExpr = $"Row.Of({string.Join(", ", cellRenders)})";
        return explicitRowIndex is { } idx ? $"{rowExpr}.AtIndex({idx})" : rowExpr;
    }

    private static string RenderSheet(Sheet sheet, string variableName)
    {
        var nextRow = 0;
        var rowRenders = sheet.Rows.Select(row => RenderRow(row, ref nextRow)).ToArray();

        var sb = new StringBuilder();
        sb.Append("var ").Append(variableName).Append(" = Sheet.Create(\n");
        sb.Append("        ").Append(RenderString(sheet.Name));
        foreach (var rowExpr in rowRenders)
            sb.Append(",\n        ").Append(rowExpr);
        sb.Append(')');

        if (sheet.MergedRanges.Count > 0)
            sb.Append("\n    .WithMergedRanges(").Append(string.Join(", ", sheet.MergedRanges.Select(RenderMergedRange))).Append(')');
        if (sheet.FreezePane is { } freezePane)
            sb.Append("\n    .WithFreezePane(").Append(RenderFreezePane(freezePane)).Append(')');
        if (sheet.AutoFilter is { } autoFilter)
            sb.Append("\n    .WithAutoFilter(").Append(RenderAutoFilterRange(autoFilter)).Append(')');
        if (sheet.Tables.Count > 0)
            sb.Append("\n    .WithTables(").Append(string.Join(", ", sheet.Tables.Select(RenderTableEntry))).Append(')');
        if (sheet.Charts.Count > 0)
            sb.Append("\n    .WithCharts(").Append(string.Join(", ", sheet.Charts.Select(RenderChartEntry))).Append(')');
        if (sheet.Images.Count > 0)
            sb.Append("\n    .WithImages(").Append(string.Join(", ", sheet.Images.Select(RenderImageEntry))).Append(')');
        if (sheet.PivotTables.Count > 0)
            sb.Append("\n    .WithPivotTables(").Append(string.Join(", ", sheet.PivotTables.Select(RenderPivotTableEntry))).Append(')');
        if (sheet.SparklineGroups.Count > 0)
            sb.Append("\n    .WithSparklineGroups(").Append(string.Join(", ", sheet.SparklineGroups.Select(RenderSparklineGroupEntry))).Append(')');
        if (sheet.ConditionalFormats.Count > 0)
            sb.Append("\n    .WithConditionalFormats(").Append(string.Join(", ", sheet.ConditionalFormats.Select(RenderConditionalFormatEntry))).Append(')');

        sb.Append(';');
        return sb.ToString();
    }

    /// <summary>
    /// Renders <paramref name="workbook"/> as a self-contained C# file that rebuilds an
    /// equivalent file at <paramref name="outputFileName"/> when run via <c>dotnet run
    /// &lt;file&gt;.cs</c> (.NET 10's "file-based apps" feature - no <c>.csproj</c> needed).
    /// <paramref name="referenceLines"/> are raw directive lines placed at the very top of
    /// the file, before anything else - typically a single <c>#:package
    /// Kookerella.CsOpenXmlDsl@X.Y.Z</c> line for a consumer of the published package, or
    /// <c>#:project ../path/to/Kookerella.CsOpenXmlDsl.csproj</c> when generating against a
    /// local build.
    /// </summary>
    public static string Generate(IReadOnlyList<string> referenceLines, string outputFileName, Workbook workbook)
    {
        var sb = new StringBuilder();

        foreach (var line in referenceLines)
            sb.Append(line).Append('\n');

        sb.Append('\n');
        sb.Append("using Kookerella.CsOpenXmlDsl;\n");
        sb.Append('\n');

        var sheetVariableNames = Enumerable.Range(0, workbook.Sheets.Count).Select(i => $"sheet{i}").ToArray();

        for (var i = 0; i < workbook.Sheets.Count; i++)
            sb.Append(RenderSheet(workbook.Sheets[i], sheetVariableNames[i])).Append('\n').Append('\n');

        var workbookExpr = $"Workbook.Create({string.Join(", ", sheetVariableNames)})";
        if (workbook.VbaProject is { } vbaProject)
            workbookExpr += $"\n    .WithVbaProject(System.Convert.FromBase64String({RenderString(Convert.ToBase64String(vbaProject))}))";

        sb.Append("var workbook = ").Append(workbookExpr).Append(";\n");
        sb.Append('\n');
        sb.Append($"WorkbookIO.Save(workbook, {RenderString(outputFileName)});\n");

        return sb.ToString();
    }
}
