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

    private static Fs.Model.MergedRange ToFsMergedRange(MergedRange range) =>
        new(
            Fs.CellRefModule.create(range.TopLeft.Row, range.TopLeft.Column),
            Fs.CellRefModule.create(range.BottomRight.Row, range.BottomRight.Column));

    private static Fs.Model.FreezePane ToFsFreezePane(FreezePane pane) => new(pane.Rows, pane.Columns);

    private static Fs.Model.AutoFilterRange ToFsAutoFilter(AutoFilterRange range) =>
        new(
            Fs.CellRefModule.create(range.TopLeft.Row, range.TopLeft.Column),
            Fs.CellRefModule.create(range.BottomRight.Row, range.BottomRight.Column));

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

        // Everything besides Cells/MergedRanges/FreezePane/AutoFilter is read straight off
        // this baseline (rather than re-derived here) so this doesn't need updating if the
        // F# core's Worksheet record ever grows another field - only the four this wrapper
        // actually models are overridden.
        var baseline = Fs.Builders.sheetOfCells(sheet.Name, ListModule.OfSeq(cells));

        return new Fs.Model.Worksheet(
            baseline.Name,
            baseline.Cells,
            baseline.ColumnProps,
            baseline.RowProps,
            ListModule.OfSeq(sheet.MergedRanges.Select(ToFsMergedRange)),
            sheet.FreezePane is { } fp ? FSharpOption<Fs.Model.FreezePane>.Some(ToFsFreezePane(fp)) : FSharpOption<Fs.Model.FreezePane>.None,
            sheet.AutoFilter is { } af ? FSharpOption<Fs.Model.AutoFilterRange>.Some(ToFsAutoFilter(af)) : FSharpOption<Fs.Model.AutoFilterRange>.None,
            baseline.Protection,
            baseline.ConditionalFormats,
            baseline.DataValidations,
            baseline.Hyperlinks,
            baseline.Comments,
            baseline.PageSetup,
            baseline.Tables,
            baseline.SparklineGroups,
            baseline.Charts,
            baseline.Images,
            baseline.PivotTables);
    }

    public static Fs.Model.Workbook ToFSharp(Workbook workbook)
    {
        // Same reasoning as ToFsWorksheet: build a baseline via the F# builder (for the
        // fields this wrapper doesn't model) and override only the ones it does, rather
        // than spelling out every field here.
        var baseline = Fs.Builders.workbook(ListModule.OfSeq(workbook.Sheets.Select(ToFsWorksheet)));

        return new Fs.Model.Workbook(
            baseline.Sheets,
            baseline.DefinedNames,
            baseline.Protection,
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
                AutoFilter = FromOption(fsSheet.AutoFilter) is { } af ? FromFsAutoFilter(af) : null
            };
        });

        return new Workbook { Sheets = sheets.ToArray(), VbaProject = FromOption(workbook.VbaProject) };
    }
}
