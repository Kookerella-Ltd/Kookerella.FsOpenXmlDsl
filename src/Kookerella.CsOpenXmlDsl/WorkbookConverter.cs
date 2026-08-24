using Microsoft.FSharp.Collections;
using Microsoft.FSharp.Core;
using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Pure translation between this wrapper's immutable C# records and the F# core's own
/// <c>Workbook</c>/<c>Worksheet</c>/<c>Cell</c>/<c>CellValue</c>/<c>CellStyle</c> values -
/// no I/O happens here at all (see <see cref="WorkbookIO"/> for the one place it does).
/// Internal: callers only ever see the C# <see cref="Workbook"/> shape, never the F# types
/// underneath. F# types are referenced via the <c>Fs</c> alias throughout rather than a
/// blanket <c>using</c>, since <c>Fs.Model.Cell</c>/<c>Fs.Model.Worksheet</c>/<c>Fs.Workbook</c> would
/// otherwise collide with this assembly's own same-named types - the same collision this
/// whole library already handles against the OOXML SDK's types (see e.g. <c>Writer.fs</c>'s
/// own doc comments on qualifying <c>Spreadsheet.Cell</c>/<c>Spreadsheet.Row</c>).
/// </summary>
internal static class WorkbookConverter
{
    private static FSharpOption<T> ToOption<T>(T? value) where T : class =>
        value is null ? FSharpOption<T>.None : FSharpOption<T>.Some(value);

    private static FSharpOption<T> ToOptionStruct<T>(T? value) where T : struct =>
        value.HasValue ? FSharpOption<T>.Some(value.Value) : FSharpOption<T>.None;

    private static T? FromOption<T>(FSharpOption<T> option) where T : class =>
        option is not null && FSharpOption<T>.get_IsSome(option) ? option.Value : null;

    private static T? FromOptionStruct<T>(FSharpOption<T> option) where T : struct =>
        option is not null && FSharpOption<T>.get_IsSome(option) ? option.Value : null;

    // ----- C# -> F# (used by WorkbookIO.Save) -----------------------------------------

    private static Fs.Styles.Color ToFsColor(RgbColor color) => Fs.Styles.Color.NewRgb(color.R, color.G, color.B);

    private static Fs.Styles.FontStyle? ToFsFont(CellStyle style)
    {
        var hasFont = style.FontName is not null
                      || style.FontSize is not null
                      || style.Bold
                      || style.Italic
                      || style.Underline
                      || style.Strikethrough
                      || style.FontColor is not null;

        if (!hasFont)
            return null;

        return new Fs.Styles.FontStyle(
            ToOption(style.FontName),
            ToOptionStruct(style.FontSize),
            style.Bold,
            style.Italic,
            style.Underline,
            style.Strikethrough,
            style.FontColor is { } fontColor ? FSharpOption<Fs.Styles.Color>.Some(ToFsColor(fontColor)) : FSharpOption<Fs.Styles.Color>.None);
    }

    private static Fs.Styles.FillStyle? ToFsFill(CellStyle style) =>
        style.FillColor is { } fillColor ? new Fs.Styles.FillStyle(ToFsColor(fillColor)) : null;

    private static Fs.Styles.BorderLineStyle ToFsBorderLineStyle(BorderLineStyle style) => style switch
    {
        BorderLineStyle.Thin => Fs.Styles.BorderLineStyle.Thin,
        BorderLineStyle.Medium => Fs.Styles.BorderLineStyle.Medium,
        BorderLineStyle.Thick => Fs.Styles.BorderLineStyle.Thick,
        BorderLineStyle.Dashed => Fs.Styles.BorderLineStyle.Dashed,
        BorderLineStyle.Dotted => Fs.Styles.BorderLineStyle.Dotted,
        BorderLineStyle.Double => Fs.Styles.BorderLineStyle.Double,
        BorderLineStyle.Hair => Fs.Styles.BorderLineStyle.Hair,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, null)
    };

    private static Fs.Styles.BorderSide ToFsBorderSide(BorderSide side) =>
        new(ToFsBorderLineStyle(side.Style), side.Color is { } c ? FSharpOption<Fs.Styles.Color>.Some(ToFsColor(c)) : FSharpOption<Fs.Styles.Color>.None);

    private static FSharpOption<Fs.Styles.BorderSide> ToFsBorderSideOption(BorderSide? side) =>
        side is { } s ? FSharpOption<Fs.Styles.BorderSide>.Some(ToFsBorderSide(s)) : FSharpOption<Fs.Styles.BorderSide>.None;

    private static Fs.Styles.BorderStyle? ToFsBorder(CellBorder? border) =>
        border is null
            ? null
            : new Fs.Styles.BorderStyle(
                ToFsBorderSideOption(border.Left),
                ToFsBorderSideOption(border.Right),
                ToFsBorderSideOption(border.Top),
                ToFsBorderSideOption(border.Bottom));

    private static Fs.Styles.HorizontalAlignment ToFsHorizontal(HorizontalCellAlignment alignment) => alignment switch
    {
        HorizontalCellAlignment.General => Fs.Styles.HorizontalAlignment.GeneralAlign,
        HorizontalCellAlignment.Left => Fs.Styles.HorizontalAlignment.AlignLeft,
        HorizontalCellAlignment.Center => Fs.Styles.HorizontalAlignment.AlignCenter,
        HorizontalCellAlignment.Right => Fs.Styles.HorizontalAlignment.AlignRight,
        HorizontalCellAlignment.Fill => Fs.Styles.HorizontalAlignment.AlignFill,
        HorizontalCellAlignment.Justify => Fs.Styles.HorizontalAlignment.AlignJustify,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null)
    };

    private static Fs.Styles.VerticalAlignment ToFsVertical(VerticalCellAlignment alignment) => alignment switch
    {
        VerticalCellAlignment.Top => Fs.Styles.VerticalAlignment.AlignTop,
        VerticalCellAlignment.Middle => Fs.Styles.VerticalAlignment.AlignMiddle,
        VerticalCellAlignment.Bottom => Fs.Styles.VerticalAlignment.AlignBottom,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null)
    };

    private static Fs.Styles.AlignmentStyle? ToFsAlignment(CellStyle style)
    {
        if (style.HorizontalAlignment is null && style.VerticalAlignment is null && !style.WrapText)
            return null;

        return new Fs.Styles.AlignmentStyle(
            style.HorizontalAlignment is { } h ? FSharpOption<Fs.Styles.HorizontalAlignment>.Some(ToFsHorizontal(h)) : FSharpOption<Fs.Styles.HorizontalAlignment>.None,
            style.VerticalAlignment is { } v ? FSharpOption<Fs.Styles.VerticalAlignment>.Some(ToFsVertical(v)) : FSharpOption<Fs.Styles.VerticalAlignment>.None,
            style.WrapText);
    }

    private static Fs.Styles.NumberFormat? ToFsNumberFormat(CellStyle style)
    {
        if (style.CustomNumberFormat is { } custom)
            return Fs.Styles.NumberFormat.NewCustom(custom);

        return style.NumberFormat switch
        {
            NumberFormatKind.General => Fs.Styles.NumberFormat.General,
            NumberFormatKind.Integer => Fs.Styles.NumberFormat.Integer,
            NumberFormatKind.TwoDecimal => Fs.Styles.NumberFormat.TwoDecimal,
            NumberFormatKind.Percentage => Fs.Styles.NumberFormat.Percentage,
            NumberFormatKind.Currency => Fs.Styles.NumberFormat.Currency,
            NumberFormatKind.ShortDate => Fs.Styles.NumberFormat.ShortDate,
            NumberFormatKind.DateAndTime => Fs.Styles.NumberFormat.DateAndTime,
            null => null,
            _ => throw new ArgumentOutOfRangeException()
        };
    }

    private static FSharpOption<Fs.Styles.CellStyle> ToFsStyle(CellStyle? style)
    {
        if (style is null)
            return FSharpOption<Fs.Styles.CellStyle>.None;

        var fsStyle = new Fs.Styles.CellStyle(
            ToOption(ToFsFont(style)),
            ToOption(ToFsFill(style)),
            ToOption(ToFsBorder(style.Border)),
            ToOption(ToFsNumberFormat(style)),
            ToOption(ToFsAlignment(style)),
            FSharpOption<Fs.Styles.CellProtection>.None);

        return FSharpOption<Fs.Styles.CellStyle>.Some(fsStyle);
    }

    private static Fs.Model.CellValue ToFsValue(CellValue value) => value switch
    {
        CellValue.Text t => Fs.Model.CellValue.NewText(t.Value),
        CellValue.Number n => Fs.Model.CellValue.NewNumber(n.Value),
        CellValue.Boolean b => Fs.Model.CellValue.NewBoolean(b.Value),
        CellValue.Date d => Fs.Model.CellValue.NewDate(d.Value),
        CellValue.Formula f => Fs.Model.CellValue.NewFormula(f.Expression, ToOptionStruct(f.CachedValue)),
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, null)
    };

    /// <summary>Assigns each cell in a row its column (explicit, or the next one after the
    /// previous cell) - a pure fold over the row's cells, mirroring the F# core's own
    /// <c>cellsForRow</c> (see <c>Builders.fs</c>) column-numbering convention exactly.</summary>
    private static IEnumerable<Fs.Model.Cell> ToFsCellsForRow(int rowIndex, Row row)
    {
        var nextColumn = 0;

        foreach (var cell in row.Cells)
        {
            var column = cell.Column ?? nextColumn;
            yield return new Fs.Model.Cell(Fs.CellRefModule.create(rowIndex, column), ToFsValue(cell.Value), ToFsStyle(cell.Style));
            nextColumn = column + 1;
        }
    }

    private static Fs.CellRef ToFsCellRef(CellPosition position) => Fs.CellRefModule.create(position.Row, position.Column);

    private static Fs.Model.MergedRange ToFsMergedRange(MergedRange range) =>
        new(ToFsCellRef(range.TopLeft), ToFsCellRef(range.BottomRight));

    private static Fs.Model.FreezePane ToFsFreezePane(FreezePane pane) => new(pane.Rows, pane.Columns);

    private static Fs.Model.AutoFilterRange ToFsAutoFilter(AutoFilterRange range) =>
        new(ToFsCellRef(range.TopLeft), ToFsCellRef(range.BottomRight));

    private static Fs.TableColumn ToFsTableColumn(TableColumn column) =>
        new(column.Name, ToOption(column.CalculatedFormula));

    private static Fs.TableStyle ToFsTableStyle(TableStyle style) =>
        new(ToOption(style.Name), style.ShowFirstColumn, style.ShowLastColumn, style.ShowRowStripes, style.ShowColumnStripes);

    private static Fs.TableEntry ToFsTableEntry(TableEntry table) =>
        new(
            ToFsCellRef(table.TopLeft),
            ToFsCellRef(table.BottomRight),
            table.Name,
            ListModule.OfSeq(table.Columns.Select(ToFsTableColumn)),
            ToFsTableStyle(table.Style));

    private static Fs.ChartType ToFsChartType(ChartType type) => type switch
    {
        ChartType.Column => Fs.ChartType.ChartColumn,
        ChartType.Bar => Fs.ChartType.ChartBar,
        ChartType.Line => Fs.ChartType.ChartLine,
        ChartType.Pie => Fs.ChartType.ChartPie,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    private static Fs.ChartSeries ToFsChartSeries(ChartSeries series) =>
        new(ToFsCellRef(series.Name), ToFsCellRef(series.ValuesTopLeft), ToFsCellRef(series.ValuesBottomRight));

    private static Fs.ChartEntry ToFsChartEntry(ChartEntry chart) =>
        new(
            ToFsChartType(chart.Type),
            ToOption(chart.Title),
            ToFsCellRef(chart.CategoriesTopLeft),
            ToFsCellRef(chart.CategoriesBottomRight),
            ListModule.OfSeq(chart.Series.Select(ToFsChartSeries)),
            chart.ShowLegend,
            ToFsCellRef(chart.TopLeftAnchor),
            ToFsCellRef(chart.BottomRightAnchor));

    private static Fs.ImageFormat ToFsImageFormat(ImageFormat format) => format switch
    {
        ImageFormat.Png => Fs.ImageFormat.Png,
        ImageFormat.Jpeg => Fs.ImageFormat.Jpeg,
        ImageFormat.Gif => Fs.ImageFormat.Gif,
        ImageFormat.Bmp => Fs.ImageFormat.Bmp,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, null)
    };

    private static Fs.ImageEntry ToFsImageEntry(ImageEntry image) =>
        new(image.Data, ToFsImageFormat(image.Format), ToFsCellRef(image.TopLeftAnchor), ToFsCellRef(image.BottomRightAnchor));

    private static Fs.PivotAggregation ToFsPivotAggregation(PivotAggregation aggregation) => aggregation switch
    {
        PivotAggregation.Sum => Fs.PivotAggregation.PivotSum,
        PivotAggregation.Count => Fs.PivotAggregation.PivotCount,
        PivotAggregation.CountNumbers => Fs.PivotAggregation.PivotCountNumbers,
        PivotAggregation.Average => Fs.PivotAggregation.PivotAverage,
        PivotAggregation.Min => Fs.PivotAggregation.PivotMin,
        PivotAggregation.Max => Fs.PivotAggregation.PivotMax,
        _ => throw new ArgumentOutOfRangeException(nameof(aggregation), aggregation, null)
    };

    private static Fs.PivotTableEntry ToFsPivotTableEntry(PivotTableEntry pivotTable) =>
        new(
            ToOption(pivotTable.SourceSheet),
            ToFsCellRef(pivotTable.SourceTopLeft),
            ToFsCellRef(pivotTable.SourceBottomRight),
            pivotTable.RowField,
            ToOption(pivotTable.ColumnField),
            pivotTable.ValueField,
            ToFsPivotAggregation(pivotTable.Aggregation),
            ToOption(pivotTable.ValueCaption),
            ToFsCellRef(pivotTable.TopLeftAnchor));

    private static Fs.SparklineType ToFsSparklineType(SparklineType type) => type switch
    {
        SparklineType.Line => Fs.SparklineType.Line,
        SparklineType.Column => Fs.SparklineType.Column,
        SparklineType.WinLoss => Fs.SparklineType.WinLoss,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    private static Fs.SparklineCell ToFsSparklineCell(SparklineCell cell) =>
        new(ToFsCellRef(cell.Cell), ToFsCellRef(cell.DataTopLeft), ToFsCellRef(cell.DataBottomRight));

    private static Fs.SparklineStyle ToFsSparklineStyle(SparklineStyle style) =>
        new(
            ToFsSparklineType(style.Type),
            style.Color is { } color ? FSharpOption<Fs.Styles.Color>.Some(ToFsColor(color)) : FSharpOption<Fs.Styles.Color>.None,
            ToOptionStruct(style.LineWeight),
            style.ShowMarkers,
            style.ShowHigh,
            style.ShowLow,
            style.ShowFirst,
            style.ShowLast,
            style.ShowNegative);

    private static Fs.SparklineGroupEntry ToFsSparklineGroupEntry(SparklineGroupEntry group) =>
        new(ToFsSparklineStyle(group.Style), ListModule.OfSeq(group.Sparklines.Select(ToFsSparklineCell)));

    /// <summary><see cref="ToFsStyle"/> always returns <c>Some</c> for a non-null input -
    /// this just unwraps that for call sites (like <see cref="ToFsConditionalFormatRule"/>)
    /// where the F# field is a plain <c>CellStyle</c>, not an <c>option</c>.</summary>
    private static Fs.Styles.CellStyle ToFsRawStyle(CellStyle style) => ToFsStyle(style).Value;

    private static Fs.ComparisonOperator ToFsComparisonOperator(ComparisonOperator op) => op switch
    {
        ComparisonOperator.Equal => Fs.ComparisonOperator.Equal,
        ComparisonOperator.NotEqual => Fs.ComparisonOperator.NotEqual,
        ComparisonOperator.GreaterThan => Fs.ComparisonOperator.GreaterThan,
        ComparisonOperator.LessThan => Fs.ComparisonOperator.LessThan,
        ComparisonOperator.GreaterThanOrEqual => Fs.ComparisonOperator.GreaterThanOrEqual,
        ComparisonOperator.LessThanOrEqual => Fs.ComparisonOperator.LessThanOrEqual,
        ComparisonOperator.Between => Fs.ComparisonOperator.Between,
        ComparisonOperator.NotBetween => Fs.ComparisonOperator.NotBetween,
        _ => throw new ArgumentOutOfRangeException(nameof(op), op, null)
    };

    private static Fs.ConditionalFormatRule ToFsConditionalFormatRule(ConditionalFormatRule rule) => rule switch
    {
        ConditionalFormatRule.CellValueRule r => Fs.ConditionalFormatRule.NewCellValueRule(
            ToFsComparisonOperator(r.Operator), r.Formula1, ToOption(r.Formula2), ToFsRawStyle(r.Style)),
        ConditionalFormatRule.FormulaRule r => Fs.ConditionalFormatRule.NewFormulaRule(r.Formula, ToFsRawStyle(r.Style)),
        ConditionalFormatRule.ColorScale2 r => Fs.ConditionalFormatRule.NewColorScale2(ToFsColor(r.MinColor), ToFsColor(r.MaxColor)),
        ConditionalFormatRule.ColorScale3 r => Fs.ConditionalFormatRule.NewColorScale3(ToFsColor(r.MinColor), ToFsColor(r.MidColor), ToFsColor(r.MaxColor)),
        ConditionalFormatRule.DataBarRule r => Fs.ConditionalFormatRule.NewDataBarRule(ToFsColor(r.Color)),
        ConditionalFormatRule.DuplicateValuesRule r => Fs.ConditionalFormatRule.NewDuplicateValuesRule(ToFsRawStyle(r.Style)),
        ConditionalFormatRule.UniqueValuesRule r => Fs.ConditionalFormatRule.NewUniqueValuesRule(ToFsRawStyle(r.Style)),
        _ => throw new ArgumentOutOfRangeException(nameof(rule), rule, null)
    };

    private static Fs.ConditionalFormatEntry ToFsConditionalFormatEntry(ConditionalFormatEntry entry) =>
        new(ToFsCellRef(entry.TopLeft), ToFsCellRef(entry.BottomRight), ToFsConditionalFormatRule(entry.Rule));

    private static Fs.ErrorAlertStyle ToFsErrorAlertStyle(ErrorAlertStyle style) => style switch
    {
        ErrorAlertStyle.Stop => Fs.ErrorAlertStyle.Stop,
        ErrorAlertStyle.Warning => Fs.ErrorAlertStyle.Warning,
        ErrorAlertStyle.Information => Fs.ErrorAlertStyle.Information,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, null)
    };

    private static Fs.ValidationAlert ToFsValidationAlert(ValidationAlert alert) =>
        new(
            alert.AllowBlank,
            ToFsErrorAlertStyle(alert.ErrorStyle),
            ToOption(alert.ErrorTitle),
            ToOption(alert.ErrorMessage),
            ToOption(alert.InputTitle),
            ToOption(alert.InputMessage));

    private static Fs.ValidationKind ToFsValidationKind(ValidationKind kind) => kind switch
    {
        ValidationKind.ListValidation k => Fs.ValidationKind.NewListValidation(ListModule.OfSeq(k.Items)),
        ValidationKind.ListFromRangeValidation k => Fs.ValidationKind.NewListFromRangeValidation(ToFsCellRef(k.TopLeft), ToFsCellRef(k.BottomRight)),
        ValidationKind.WholeNumberValidation k => Fs.ValidationKind.NewWholeNumberValidation(ToFsComparisonOperator(k.Operator), k.Formula1, ToOption(k.Formula2)),
        ValidationKind.DecimalValidation k => Fs.ValidationKind.NewDecimalValidation(ToFsComparisonOperator(k.Operator), k.Formula1, ToOption(k.Formula2)),
        ValidationKind.TextLengthValidation k => Fs.ValidationKind.NewTextLengthValidation(ToFsComparisonOperator(k.Operator), k.Formula1, ToOption(k.Formula2)),
        ValidationKind.CustomValidation k => Fs.ValidationKind.NewCustomValidation(k.Formula),
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
    };

    private static Fs.DataValidationEntry ToFsDataValidationEntry(DataValidationEntry entry) =>
        new(ToFsCellRef(entry.TopLeft), ToFsCellRef(entry.BottomRight), ToFsValidationKind(entry.Kind), ToFsValidationAlert(entry.Alert));

    private static Fs.HyperlinkTarget ToFsHyperlinkTarget(HyperlinkTarget target) => target switch
    {
        HyperlinkTarget.ExternalHyperlink t => Fs.HyperlinkTarget.NewExternalHyperlink(t.Url),
        HyperlinkTarget.InternalHyperlink t => Fs.HyperlinkTarget.NewInternalHyperlink(t.Location),
        _ => throw new ArgumentOutOfRangeException(nameof(target), target, null)
    };

    private static Fs.HyperlinkEntry ToFsHyperlinkEntry(HyperlinkEntry entry) =>
        new(ToFsCellRef(entry.TopLeft), ToFsCellRef(entry.BottomRight), ToFsHyperlinkTarget(entry.Target), ToOption(entry.Tooltip), ToOption(entry.Display));

    private static Fs.CommentEntry ToFsCommentEntry(CommentEntry entry) => new(ToFsCellRef(entry.Cell), entry.Author, entry.Text);

    private static Fs.PageOrientation ToFsPageOrientation(PageOrientation orientation) => orientation switch
    {
        PageOrientation.Portrait => Fs.PageOrientation.Portrait,
        PageOrientation.Landscape => Fs.PageOrientation.Landscape,
        _ => throw new ArgumentOutOfRangeException(nameof(orientation), orientation, null)
    };

    private static Fs.PaperSize ToFsPaperSize(PaperSize size) => size switch
    {
        PaperSize.Letter => Fs.PaperSize.Letter,
        PaperSize.Legal => Fs.PaperSize.Legal,
        PaperSize.Tabloid => Fs.PaperSize.Tabloid,
        PaperSize.A3 => Fs.PaperSize.A3,
        PaperSize.A4 => Fs.PaperSize.A4,
        PaperSize.OtherPaperSize s => Fs.PaperSize.NewOtherPaperSize(s.Code),
        _ => throw new ArgumentOutOfRangeException(nameof(size), size, null)
    };

    private static Fs.PrintScaling ToFsPrintScaling(PrintScaling scaling) => scaling switch
    {
        PrintScaling.ScalePercent s => Fs.PrintScaling.NewScalePercent(s.Percent),
        PrintScaling.FitToPage s => Fs.PrintScaling.NewFitToPage(s.Width, s.Height),
        _ => throw new ArgumentOutOfRangeException(nameof(scaling), scaling, null)
    };

    private static Fs.PageMargins ToFsPageMargins(PageMargins margins) =>
        new(margins.Left, margins.Right, margins.Top, margins.Bottom, margins.Header, margins.Footer);

    private static Fs.PageSetup ToFsPageSetup(PageSetup pageSetup) =>
        new(
            ToFsPageOrientation(pageSetup.Orientation),
            pageSetup.PaperSize is { } paperSize ? FSharpOption<Fs.PaperSize>.Some(ToFsPaperSize(paperSize)) : FSharpOption<Fs.PaperSize>.None,
            pageSetup.Scaling is { } scaling ? FSharpOption<Fs.PrintScaling>.Some(ToFsPrintScaling(scaling)) : FSharpOption<Fs.PrintScaling>.None,
            ToFsPageMargins(pageSetup.Margins),
            ListModule.OfSeq(pageSetup.PrintArea.Select(r => Tuple.Create(ToFsCellRef(r.TopLeft), ToFsCellRef(r.BottomRight)))),
            ToOption(pageSetup.Header),
            ToOption(pageSetup.Footer),
            ToOption(pageSetup.EvenHeader),
            ToOption(pageSetup.EvenFooter),
            ToOption(pageSetup.FirstHeader),
            ToOption(pageSetup.FirstFooter));

    private static Fs.SheetProtection ToFsSheetProtection(SheetProtection protection) =>
        new(
            ToOption(protection.Password),
            protection.Sheet,
            ToOptionStruct(protection.ObjectsBlocked),
            ToOptionStruct(protection.ScenariosBlocked),
            ToOptionStruct(protection.FormatCellsBlocked),
            ToOptionStruct(protection.FormatColumnsBlocked),
            ToOptionStruct(protection.FormatRowsBlocked),
            ToOptionStruct(protection.InsertColumnsBlocked),
            ToOptionStruct(protection.InsertRowsBlocked),
            ToOptionStruct(protection.InsertHyperlinksBlocked),
            ToOptionStruct(protection.DeleteColumnsBlocked),
            ToOptionStruct(protection.DeleteRowsBlocked),
            ToOptionStruct(protection.SelectLockedCellsBlocked),
            ToOptionStruct(protection.SortBlocked),
            ToOptionStruct(protection.AutoFilterBlocked),
            ToOptionStruct(protection.PivotTablesBlocked),
            ToOptionStruct(protection.SelectUnlockedCellsBlocked));

    private static Fs.Model.Worksheet ToFsWorksheet(Sheet sheet)
    {
        var nextRow = 0;
        var cells = new List<Fs.Model.Cell>();

        foreach (var row in sheet.Rows)
        {
            var rowIndex = row.Index ?? nextRow;
            cells.AddRange(ToFsCellsForRow(rowIndex, row));
            nextRow = rowIndex + 1;
        }

        // Everything besides Cells/MergedRanges/FreezePane/AutoFilter/Tables is read
        // straight off this baseline (rather than re-derived here) so this doesn't need
        // updating if the F# core's Worksheet record ever grows another field - only the
        // fields this wrapper actually models are overridden.
        var baseline = Fs.Builders.sheetOfCells(sheet.Name, ListModule.OfSeq(cells));

        return new Fs.Model.Worksheet(
            baseline.Name,
            baseline.Cells,
            baseline.ColumnProps,
            baseline.RowProps,
            ListModule.OfSeq(sheet.MergedRanges.Select(ToFsMergedRange)),
            sheet.FreezePane is { } fp ? FSharpOption<Fs.Model.FreezePane>.Some(ToFsFreezePane(fp)) : FSharpOption<Fs.Model.FreezePane>.None,
            sheet.AutoFilter is { } af ? FSharpOption<Fs.Model.AutoFilterRange>.Some(ToFsAutoFilter(af)) : FSharpOption<Fs.Model.AutoFilterRange>.None,
            sheet.Protection is { } protection ? FSharpOption<Fs.SheetProtection>.Some(ToFsSheetProtection(protection)) : FSharpOption<Fs.SheetProtection>.None,
            ListModule.OfSeq(sheet.ConditionalFormats.Select(ToFsConditionalFormatEntry)),
            ListModule.OfSeq(sheet.DataValidations.Select(ToFsDataValidationEntry)),
            ListModule.OfSeq(sheet.Hyperlinks.Select(ToFsHyperlinkEntry)),
            ListModule.OfSeq(sheet.Comments.Select(ToFsCommentEntry)),
            sheet.PageSetup is { } pageSetup ? FSharpOption<Fs.PageSetup>.Some(ToFsPageSetup(pageSetup)) : FSharpOption<Fs.PageSetup>.None,
            ListModule.OfSeq(sheet.Tables.Select(ToFsTableEntry)),
            ListModule.OfSeq(sheet.SparklineGroups.Select(ToFsSparklineGroupEntry)),
            ListModule.OfSeq(sheet.Charts.Select(ToFsChartEntry)),
            ListModule.OfSeq(sheet.Images.Select(ToFsImageEntry)),
            ListModule.OfSeq(sheet.PivotTables.Select(ToFsPivotTableEntry)));
    }

    private static Fs.DefinedNameScope ToFsDefinedNameScope(DefinedNameScope scope) => scope switch
    {
        DefinedNameScope.WorkbookScope => Fs.DefinedNameScope.WorkbookScope,
        DefinedNameScope.SheetScope s => Fs.DefinedNameScope.NewSheetScope(s.SheetName),
        _ => throw new ArgumentOutOfRangeException(nameof(scope), scope, null)
    };

    private static Fs.DefinedNameEntry ToFsDefinedNameEntry(DefinedNameEntry entry) =>
        new(entry.Name, entry.Formula, ToFsDefinedNameScope(entry.Scope), entry.Hidden);

    private static Fs.WorkbookProtection ToFsWorkbookProtection(WorkbookProtection protection) =>
        new(ToOption(protection.Password), ToOptionStruct(protection.LockStructure), ToOptionStruct(protection.LockWindows));

    public static Fs.Model.Workbook ToFSharp(Workbook workbook)
    {
        // Same reasoning as ToFsWorksheet: build a baseline via the F# builder (for the
        // fields this wrapper doesn't model) and override only the ones it does, rather
        // than spelling out every field here.
        var baseline = Fs.Builders.workbook(ListModule.OfSeq(workbook.Sheets.Select(ToFsWorksheet)));

        return new Fs.Model.Workbook(
            baseline.Sheets,
            ListModule.OfSeq(workbook.DefinedNames.Select(ToFsDefinedNameEntry)),
            workbook.Protection is { } protection ? FSharpOption<Fs.WorkbookProtection>.Some(ToFsWorkbookProtection(protection)) : FSharpOption<Fs.WorkbookProtection>.None,
            ToOption(workbook.VbaProject));
    }

    // ----- F# -> C# (used by WorkbookIO.Load) -----------------------------------------

    private static RgbColor? FromFsColor(Fs.Styles.Color color) => color switch
    {
        Fs.Styles.Color.Rgb rgb => new RgbColor(rgb.red, rgb.green, rgb.blue),
        // Indexed/Theme colors aren't resolved to RGB by the F# core itself (it doesn't
        // parse the workbook's theme part) - this wrapper can't show a meaningful color
        // for them either, so they come back as no explicit color rather than a guess.
        _ => null
    };

    private static void ApplyFsFont(Fs.Styles.FontStyle font, ref CellStyle style)
    {
        style = style with
        {
            FontName = FromOption(font.Name),
            FontSize = FromOptionStruct(font.Size),
            Bold = font.Bold,
            Italic = font.Italic,
            Underline = font.Underline,
            Strikethrough = font.Strikethrough,
            FontColor = font.Color is { } c && FromOption(c) is { } fc ? FromFsColor(fc) : style.FontColor
        };
    }

    private static BorderSide? FromFsBorderSide(FSharpOption<Fs.Styles.BorderSide> option)
    {
        if (FromOption(option) is not { } side)
            return null;

        var lineStyle = side.Style switch
        {
            Fs.Styles.BorderLineStyle style when style.IsThin => BorderLineStyle.Thin,
            Fs.Styles.BorderLineStyle style when style.IsMedium => BorderLineStyle.Medium,
            Fs.Styles.BorderLineStyle style when style.IsThick => BorderLineStyle.Thick,
            Fs.Styles.BorderLineStyle style when style.IsDashed => BorderLineStyle.Dashed,
            Fs.Styles.BorderLineStyle style when style.IsDotted => BorderLineStyle.Dotted,
            Fs.Styles.BorderLineStyle style when style.IsDouble => BorderLineStyle.Double,
            Fs.Styles.BorderLineStyle style when style.IsHair => BorderLineStyle.Hair,
            // `Other` (a raw OOXML style name Core doesn't have a named case for) has no
            // equivalent in this wrapper's narrower enum - dropped rather than guessed at.
            _ => (BorderLineStyle?)null
        };

        if (lineStyle is null)
            return null;

        var color = FromOption(side.Color);
        return new BorderSide(lineStyle.Value, color is { } c ? FromFsColor(c) : null);
    }

    private static CellStyle ApplyFsStyle(FSharpOption<Fs.Styles.CellStyle> fsStyleOption)
    {
        var style = CellStyle.Default;

        if (FromOption(fsStyleOption) is not { } fsStyle)
            return style;

        if (FromOption(fsStyle.Font) is { } font)
            ApplyFsFont(font, ref style);

        if (FromOption(fsStyle.Fill) is { } fill)
            style = style with { FillColor = FromFsColor(fill.Color) };

        if (FromOption(fsStyle.Border) is { } border)
        {
            var csBorder = CellBorder.None with
            {
                Left = FromFsBorderSide(border.Left),
                Right = FromFsBorderSide(border.Right),
                Top = FromFsBorderSide(border.Top),
                Bottom = FromFsBorderSide(border.Bottom)
            };
            style = style with { Border = csBorder };
        }

        if (FromOption(fsStyle.Alignment) is { } alignment)
        {
            var horizontal = FromOption(alignment.Horizontal) switch
            {
                { } h when h.IsGeneralAlign => HorizontalCellAlignment.General,
                { } h when h.IsAlignLeft => HorizontalCellAlignment.Left,
                { } h when h.IsAlignCenter => HorizontalCellAlignment.Center,
                { } h when h.IsAlignRight => HorizontalCellAlignment.Right,
                { } h when h.IsAlignFill => HorizontalCellAlignment.Fill,
                { } h when h.IsAlignJustify => HorizontalCellAlignment.Justify,
                _ => (HorizontalCellAlignment?)null
            };

            var vertical = FromOption(alignment.Vertical) switch
            {
                { } v when v.IsAlignTop => VerticalCellAlignment.Top,
                { } v when v.IsAlignMiddle => VerticalCellAlignment.Middle,
                { } v when v.IsAlignBottom => VerticalCellAlignment.Bottom,
                _ => (VerticalCellAlignment?)null
            };

            style = style with { HorizontalAlignment = horizontal, VerticalAlignment = vertical, WrapText = alignment.WrapText };
        }

        if (FromOption(fsStyle.NumberFormat) is { } numberFormat)
        {
            style = numberFormat switch
            {
                Fs.Styles.NumberFormat nf when nf.IsGeneral => style with { NumberFormat = NumberFormatKind.General },
                Fs.Styles.NumberFormat nf when nf.IsInteger => style with { NumberFormat = NumberFormatKind.Integer },
                Fs.Styles.NumberFormat nf when nf.IsTwoDecimal => style with { NumberFormat = NumberFormatKind.TwoDecimal },
                Fs.Styles.NumberFormat nf when nf.IsPercentage => style with { NumberFormat = NumberFormatKind.Percentage },
                Fs.Styles.NumberFormat nf when nf.IsCurrency => style with { NumberFormat = NumberFormatKind.Currency },
                Fs.Styles.NumberFormat nf when nf.IsShortDate => style with { NumberFormat = NumberFormatKind.ShortDate },
                Fs.Styles.NumberFormat nf when nf.IsDateAndTime => style with { NumberFormat = NumberFormatKind.DateAndTime },
                Fs.Styles.NumberFormat.Custom custom => style with { CustomNumberFormat = custom.formatCode },
                _ => style
            };
        }

        return style;
    }

    private static CellValue? FromFsValue(Fs.Model.CellValue value) => value switch
    {
        Fs.Model.CellValue.Text t => new CellValue.Text(t.Item),
        Fs.Model.CellValue.Number n => new CellValue.Number(n.Item),
        Fs.Model.CellValue.Boolean b => new CellValue.Boolean(b.Item),
        Fs.Model.CellValue.Date d => new CellValue.Date(d.Item),
        Fs.Model.CellValue.Formula f => new CellValue.Formula(f.expression, FromOptionStruct(f.cachedValue)),
        // `Empty` has no equivalent case in this wrapper's CellValue - see its own doc
        // comment on why (just don't add a cell with nothing in it); dropped on read.
        _ => null
    };

    private static CellPosition FromFsCellRef(Fs.CellRef cellRef) => new(cellRef.Row, cellRef.Col);

    private static MergedRange FromFsMergedRange(Fs.Model.MergedRange range) =>
        new(FromFsCellRef(range.TopLeft), FromFsCellRef(range.BottomRight));

    private static FreezePane FromFsFreezePane(Fs.Model.FreezePane pane) => new(pane.Rows, pane.Columns);

    private static AutoFilterRange FromFsAutoFilter(Fs.Model.AutoFilterRange range) =>
        new(FromFsCellRef(range.TopLeft), FromFsCellRef(range.BottomRight));

    private static TableColumn FromFsTableColumn(Fs.TableColumn column) =>
        new(column.Name, FromOption(column.CalculatedFormula));

    private static TableStyle FromFsTableStyle(Fs.TableStyle style) => new()
    {
        Name = FromOption(style.Name),
        ShowFirstColumn = style.ShowFirstColumn,
        ShowLastColumn = style.ShowLastColumn,
        ShowRowStripes = style.ShowRowStripes,
        ShowColumnStripes = style.ShowColumnStripes
    };

    private static TableEntry FromFsTableEntry(Fs.TableEntry table) =>
        new(
            FromFsCellRef(table.TopLeft),
            FromFsCellRef(table.BottomRight),
            table.Name,
            table.Columns.Select(FromFsTableColumn).ToArray())
        {
            Style = FromFsTableStyle(table.Style)
        };

    private static ChartType FromFsChartType(Fs.ChartType type) => type switch
    {
        { IsChartColumn: true } => ChartType.Column,
        { IsChartBar: true } => ChartType.Bar,
        { IsChartLine: true } => ChartType.Line,
        { IsChartPie: true } => ChartType.Pie,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    private static ChartSeries FromFsChartSeries(Fs.ChartSeries series) =>
        new(FromFsCellRef(series.Name), FromFsCellRef(series.ValuesTopLeft), FromFsCellRef(series.ValuesBottomRight));

    private static ChartEntry FromFsChartEntry(Fs.ChartEntry chart) =>
        new(
            FromFsChartType(chart.Type),
            FromFsCellRef(chart.CategoriesTopLeft),
            FromFsCellRef(chart.CategoriesBottomRight),
            FromFsCellRef(chart.TopLeftAnchor),
            FromFsCellRef(chart.BottomRightAnchor),
            chart.Series.Select(FromFsChartSeries).ToArray())
        {
            Title = FromOption(chart.Title),
            ShowLegend = chart.ShowLegend
        };

    private static ImageFormat FromFsImageFormat(Fs.ImageFormat format) => format switch
    {
        { IsPng: true } => ImageFormat.Png,
        { IsJpeg: true } => ImageFormat.Jpeg,
        { IsGif: true } => ImageFormat.Gif,
        { IsBmp: true } => ImageFormat.Bmp,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, null)
    };

    private static ImageEntry FromFsImageEntry(Fs.ImageEntry image) =>
        new(image.Data, FromFsImageFormat(image.Format), FromFsCellRef(image.TopLeftAnchor), FromFsCellRef(image.BottomRightAnchor));

    private static PivotAggregation FromFsPivotAggregation(Fs.PivotAggregation aggregation) => aggregation switch
    {
        { IsPivotSum: true } => PivotAggregation.Sum,
        { IsPivotCount: true } => PivotAggregation.Count,
        { IsPivotCountNumbers: true } => PivotAggregation.CountNumbers,
        { IsPivotAverage: true } => PivotAggregation.Average,
        { IsPivotMin: true } => PivotAggregation.Min,
        { IsPivotMax: true } => PivotAggregation.Max,
        _ => throw new ArgumentOutOfRangeException(nameof(aggregation), aggregation, null)
    };

    private static PivotTableEntry FromFsPivotTableEntry(Fs.PivotTableEntry pivotTable) =>
        new(
            FromFsCellRef(pivotTable.SourceTopLeft),
            FromFsCellRef(pivotTable.SourceBottomRight),
            pivotTable.RowField,
            pivotTable.ValueField,
            FromFsCellRef(pivotTable.TopLeftAnchor))
        {
            SourceSheet = FromOption(pivotTable.SourceSheet),
            ColumnField = FromOption(pivotTable.ColumnField),
            Aggregation = FromFsPivotAggregation(pivotTable.Aggregation),
            ValueCaption = FromOption(pivotTable.ValueCaption)
        };

    private static SparklineType FromFsSparklineType(Fs.SparklineType type) => type switch
    {
        { IsLine: true } => SparklineType.Line,
        { IsColumn: true } => SparklineType.Column,
        { IsWinLoss: true } => SparklineType.WinLoss,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    private static SparklineCell FromFsSparklineCell(Fs.SparklineCell cell) =>
        new(FromFsCellRef(cell.Cell), FromFsCellRef(cell.DataTopLeft), FromFsCellRef(cell.DataBottomRight));

    private static SparklineStyle FromFsSparklineStyle(Fs.SparklineStyle style) => new()
    {
        Type = FromFsSparklineType(style.Type),
        Color = FromOption(style.Color) is { } color ? FromFsColor(color) : null,
        LineWeight = FromOptionStruct(style.LineWeight),
        ShowMarkers = style.ShowMarkers,
        ShowHigh = style.ShowHigh,
        ShowLow = style.ShowLow,
        ShowFirst = style.ShowFirst,
        ShowLast = style.ShowLast,
        ShowNegative = style.ShowNegative
    };

    private static SparklineGroupEntry FromFsSparklineGroupEntry(Fs.SparklineGroupEntry group) =>
        new(group.Sparklines.Select(FromFsSparklineCell).ToArray()) { Style = FromFsSparklineStyle(group.Style) };

    /// <summary>Wraps a raw (non-<c>option</c>) F# <c>CellStyle</c> - the shape
    /// <see cref="ConditionalFormatRule"/>'s cases carry - so it can reuse <see
    /// cref="ApplyFsStyle"/>'s translation logic rather than duplicating it.</summary>
    private static CellStyle ApplyFsRawStyle(Fs.Styles.CellStyle style) => ApplyFsStyle(FSharpOption<Fs.Styles.CellStyle>.Some(style));

    private static ComparisonOperator FromFsComparisonOperator(Fs.ComparisonOperator op) => op switch
    {
        { IsEqual: true } => ComparisonOperator.Equal,
        { IsNotEqual: true } => ComparisonOperator.NotEqual,
        { IsGreaterThan: true } => ComparisonOperator.GreaterThan,
        { IsLessThan: true } => ComparisonOperator.LessThan,
        { IsGreaterThanOrEqual: true } => ComparisonOperator.GreaterThanOrEqual,
        { IsLessThanOrEqual: true } => ComparisonOperator.LessThanOrEqual,
        { IsBetween: true } => ComparisonOperator.Between,
        { IsNotBetween: true } => ComparisonOperator.NotBetween,
        _ => throw new ArgumentOutOfRangeException(nameof(op), op, null)
    };

    private static ConditionalFormatRule FromFsConditionalFormatRule(Fs.ConditionalFormatRule rule) => rule switch
    {
        Fs.ConditionalFormatRule.CellValueRule r =>
            new ConditionalFormatRule.CellValueRule(FromFsComparisonOperator(r.@operator), r.formula1, FromOption(r.formula2), ApplyFsRawStyle(r.style)),
        Fs.ConditionalFormatRule.FormulaRule r => new ConditionalFormatRule.FormulaRule(r.formula, ApplyFsRawStyle(r.style)),
        Fs.ConditionalFormatRule.ColorScale2 r =>
            new ConditionalFormatRule.ColorScale2(FromFsColor(r.minColor) ?? RgbColor.Black, FromFsColor(r.maxColor) ?? RgbColor.Black),
        Fs.ConditionalFormatRule.ColorScale3 r =>
            new ConditionalFormatRule.ColorScale3(FromFsColor(r.minColor) ?? RgbColor.Black, FromFsColor(r.midColor) ?? RgbColor.Black, FromFsColor(r.maxColor) ?? RgbColor.Black),
        Fs.ConditionalFormatRule.DataBarRule r => new ConditionalFormatRule.DataBarRule(FromFsColor(r.color) ?? RgbColor.Black),
        Fs.ConditionalFormatRule.DuplicateValuesRule r => new ConditionalFormatRule.DuplicateValuesRule(ApplyFsRawStyle(r.style)),
        Fs.ConditionalFormatRule.UniqueValuesRule r => new ConditionalFormatRule.UniqueValuesRule(ApplyFsRawStyle(r.style)),
        _ => throw new ArgumentOutOfRangeException(nameof(rule), rule, null)
    };

    private static ConditionalFormatEntry FromFsConditionalFormatEntry(Fs.ConditionalFormatEntry entry) =>
        new(FromFsCellRef(entry.TopLeft), FromFsCellRef(entry.BottomRight), FromFsConditionalFormatRule(entry.Rule));

    private static ErrorAlertStyle FromFsErrorAlertStyle(Fs.ErrorAlertStyle style) => style switch
    {
        { IsStop: true } => ErrorAlertStyle.Stop,
        { IsWarning: true } => ErrorAlertStyle.Warning,
        { IsInformation: true } => ErrorAlertStyle.Information,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, null)
    };

    private static ValidationAlert FromFsValidationAlert(Fs.ValidationAlert alert) => new()
    {
        AllowBlank = alert.AllowBlank,
        ErrorStyle = FromFsErrorAlertStyle(alert.ErrorStyle),
        ErrorTitle = FromOption(alert.ErrorTitle),
        ErrorMessage = FromOption(alert.ErrorMessage),
        InputTitle = FromOption(alert.InputTitle),
        InputMessage = FromOption(alert.InputMessage)
    };

    private static ValidationKind FromFsValidationKind(Fs.ValidationKind kind) => kind switch
    {
        Fs.ValidationKind.ListValidation k => new ValidationKind.ListValidation(k.items.ToArray()),
        Fs.ValidationKind.ListFromRangeValidation k => new ValidationKind.ListFromRangeValidation(FromFsCellRef(k.topLeft), FromFsCellRef(k.bottomRight)),
        Fs.ValidationKind.WholeNumberValidation k => new ValidationKind.WholeNumberValidation(FromFsComparisonOperator(k.@operator), k.formula1, FromOption(k.formula2)),
        Fs.ValidationKind.DecimalValidation k => new ValidationKind.DecimalValidation(FromFsComparisonOperator(k.@operator), k.formula1, FromOption(k.formula2)),
        Fs.ValidationKind.TextLengthValidation k => new ValidationKind.TextLengthValidation(FromFsComparisonOperator(k.@operator), k.formula1, FromOption(k.formula2)),
        Fs.ValidationKind.CustomValidation k => new ValidationKind.CustomValidation(k.formula),
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
    };

    private static DataValidationEntry FromFsDataValidationEntry(Fs.DataValidationEntry entry) =>
        new(FromFsCellRef(entry.TopLeft), FromFsCellRef(entry.BottomRight), FromFsValidationKind(entry.Kind))
        {
            Alert = FromFsValidationAlert(entry.Alert)
        };

    private static HyperlinkTarget FromFsHyperlinkTarget(Fs.HyperlinkTarget target) => target switch
    {
        Fs.HyperlinkTarget.ExternalHyperlink t => new HyperlinkTarget.ExternalHyperlink(t.url),
        Fs.HyperlinkTarget.InternalHyperlink t => new HyperlinkTarget.InternalHyperlink(t.location),
        _ => throw new ArgumentOutOfRangeException(nameof(target), target, null)
    };

    private static HyperlinkEntry FromFsHyperlinkEntry(Fs.HyperlinkEntry entry) =>
        new(FromFsCellRef(entry.TopLeft), FromFsCellRef(entry.BottomRight), FromFsHyperlinkTarget(entry.Target))
        {
            Tooltip = FromOption(entry.Tooltip),
            Display = FromOption(entry.Display)
        };

    private static CommentEntry FromFsCommentEntry(Fs.CommentEntry entry) => new(FromFsCellRef(entry.Cell), entry.Text, entry.Author);

    private static PageOrientation FromFsPageOrientation(Fs.PageOrientation orientation) => orientation switch
    {
        { IsPortrait: true } => PageOrientation.Portrait,
        { IsLandscape: true } => PageOrientation.Landscape,
        _ => throw new ArgumentOutOfRangeException(nameof(orientation), orientation, null)
    };

    private static PaperSize FromFsPaperSize(Fs.PaperSize size) => size switch
    {
        { IsLetter: true } => new PaperSize.Letter(),
        { IsLegal: true } => new PaperSize.Legal(),
        { IsTabloid: true } => new PaperSize.Tabloid(),
        { IsA3: true } => new PaperSize.A3(),
        { IsA4: true } => new PaperSize.A4(),
        Fs.PaperSize.OtherPaperSize s => new PaperSize.OtherPaperSize(s.code),
        _ => throw new ArgumentOutOfRangeException(nameof(size), size, null)
    };

    private static PrintScaling FromFsPrintScaling(Fs.PrintScaling scaling) => scaling switch
    {
        Fs.PrintScaling.ScalePercent s => new PrintScaling.ScalePercent(s.percent),
        Fs.PrintScaling.FitToPage s => new PrintScaling.FitToPage(s.width, s.height),
        _ => throw new ArgumentOutOfRangeException(nameof(scaling), scaling, null)
    };

    private static PageMargins FromFsPageMargins(Fs.PageMargins margins) => new()
    {
        Left = margins.Left,
        Right = margins.Right,
        Top = margins.Top,
        Bottom = margins.Bottom,
        Header = margins.Header,
        Footer = margins.Footer
    };

    private static PageSetup FromFsPageSetup(Fs.PageSetup pageSetup) => new()
    {
        Orientation = FromFsPageOrientation(pageSetup.Orientation),
        PaperSize = FromOption(pageSetup.PaperSize) is { } paperSize ? FromFsPaperSize(paperSize) : null,
        Scaling = FromOption(pageSetup.Scaling) is { } scaling ? FromFsPrintScaling(scaling) : null,
        Margins = FromFsPageMargins(pageSetup.Margins),
        PrintArea = pageSetup.PrintArea.Select(r => (FromFsCellRef(r.Item1), FromFsCellRef(r.Item2))).ToArray(),
        Header = FromOption(pageSetup.Header),
        Footer = FromOption(pageSetup.Footer),
        EvenHeader = FromOption(pageSetup.EvenHeader),
        EvenFooter = FromOption(pageSetup.EvenFooter),
        FirstHeader = FromOption(pageSetup.FirstHeader),
        FirstFooter = FromOption(pageSetup.FirstFooter)
    };

    private static SheetProtection FromFsSheetProtection(Fs.SheetProtection protection) => new()
    {
        Password = FromOption(protection.Password),
        Sheet = protection.Sheet,
        ObjectsBlocked = FromOptionStruct(protection.Objects),
        ScenariosBlocked = FromOptionStruct(protection.Scenarios),
        FormatCellsBlocked = FromOptionStruct(protection.FormatCells),
        FormatColumnsBlocked = FromOptionStruct(protection.FormatColumns),
        FormatRowsBlocked = FromOptionStruct(protection.FormatRows),
        InsertColumnsBlocked = FromOptionStruct(protection.InsertColumns),
        InsertRowsBlocked = FromOptionStruct(protection.InsertRows),
        InsertHyperlinksBlocked = FromOptionStruct(protection.InsertHyperlinks),
        DeleteColumnsBlocked = FromOptionStruct(protection.DeleteColumns),
        DeleteRowsBlocked = FromOptionStruct(protection.DeleteRows),
        SelectLockedCellsBlocked = FromOptionStruct(protection.SelectLockedCells),
        SortBlocked = FromOptionStruct(protection.Sort),
        AutoFilterBlocked = FromOptionStruct(protection.AutoFilter),
        PivotTablesBlocked = FromOptionStruct(protection.PivotTables),
        SelectUnlockedCellsBlocked = FromOptionStruct(protection.SelectUnlockedCells)
    };

    private static WorkbookProtection FromFsWorkbookProtection(Fs.WorkbookProtection protection) => new()
    {
        Password = FromOption(protection.Password),
        LockStructure = FromOptionStruct(protection.LockStructure),
        LockWindows = FromOptionStruct(protection.LockWindows)
    };

    private static DefinedNameScope FromFsDefinedNameScope(Fs.DefinedNameScope scope) => scope switch
    {
        { IsWorkbookScope: true } => new DefinedNameScope.WorkbookScope(),
        Fs.DefinedNameScope.SheetScope s => new DefinedNameScope.SheetScope(s.sheetName),
        _ => throw new ArgumentOutOfRangeException(nameof(scope), scope, null)
    };

    private static DefinedNameEntry FromFsDefinedNameEntry(Fs.DefinedNameEntry entry) =>
        new(entry.Name, entry.Formula, FromFsDefinedNameScope(entry.Scope), entry.Hidden);

    public static Workbook FromFSharp(Fs.Model.Workbook workbook)
    {
        var sheets = workbook.Sheets.Select(fsSheet =>
        {
            var rowsByIndex = new SortedDictionary<int, List<Cell>>();

            foreach (var fsCell in fsSheet.Cells)
            {
                if (FromFsValue(fsCell.Value) is not { } value)
                    continue;

                if (!rowsByIndex.TryGetValue(fsCell.Ref.Row, out var rowCells))
                {
                    rowCells = [];
                    rowsByIndex[fsCell.Ref.Row] = rowCells;
                }

                var style = ApplyFsStyle(fsCell.Style);
                var cell = new Cell(value) { Column = fsCell.Ref.Col, Style = style == CellStyle.Default ? null : style };
                rowCells.Add(cell);
            }

            var rows = rowsByIndex.Select(kvp => new Row { Index = kvp.Key, Cells = kvp.Value }).ToArray();

            return new Sheet(fsSheet.Name)
            {
                Rows = rows,
                MergedRanges = fsSheet.MergedRanges.Select(FromFsMergedRange).ToArray(),
                FreezePane = FromOption(fsSheet.FreezePane) is { } fp ? FromFsFreezePane(fp) : null,
                AutoFilter = FromOption(fsSheet.AutoFilter) is { } af ? FromFsAutoFilter(af) : null,
                Tables = fsSheet.Tables.Select(FromFsTableEntry).ToArray(),
                Charts = fsSheet.Charts.Select(FromFsChartEntry).ToArray(),
                Images = fsSheet.Images.Select(FromFsImageEntry).ToArray(),
                PivotTables = fsSheet.PivotTables.Select(FromFsPivotTableEntry).ToArray(),
                SparklineGroups = fsSheet.SparklineGroups.Select(FromFsSparklineGroupEntry).ToArray(),
                ConditionalFormats = fsSheet.ConditionalFormats.Select(FromFsConditionalFormatEntry).ToArray(),
                DataValidations = fsSheet.DataValidations.Select(FromFsDataValidationEntry).ToArray(),
                Hyperlinks = fsSheet.Hyperlinks.Select(FromFsHyperlinkEntry).ToArray(),
                Comments = fsSheet.Comments.Select(FromFsCommentEntry).ToArray(),
                PageSetup = FromOption(fsSheet.PageSetup) is { } pageSetup ? FromFsPageSetup(pageSetup) : null,
                Protection = FromOption(fsSheet.Protection) is { } sheetProtection ? FromFsSheetProtection(sheetProtection) : null
            };
        });

        return new Workbook
        {
            Sheets = sheets.ToArray(),
            DefinedNames = workbook.DefinedNames.Select(FromFsDefinedNameEntry).ToArray(),
            Protection = FromOption(workbook.Protection) is { } workbookProtection ? FromFsWorkbookProtection(workbookProtection) : null,
            VbaProject = FromOption(workbook.VbaProject)
        };
    }
}
