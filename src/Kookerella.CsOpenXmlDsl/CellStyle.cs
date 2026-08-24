namespace Kookerella.CsOpenXmlDsl;

public enum HorizontalCellAlignment
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify
}

public enum VerticalCellAlignment
{
    Top,
    Middle,
    Bottom
}

/// <summary>Covers the common OOXML border line styles - matches the F# core's named
/// cases exactly (its <c>Other</c> escape hatch for uncommon styles isn't exposed at this
/// layer).</summary>
public enum BorderLineStyle
{
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair
}

/// <summary>One border edge - a line style plus an optional color (<see langword="null"/>
/// means Excel's own default border color, usually black).</summary>
public sealed record BorderSide(BorderLineStyle Style, RgbColor? Color = null);

/// <summary>The four edges of a cell border, each independently optional. Diagonal borders
/// aren't modeled - see the F# core's own <c>BorderStyle</c> doc comment.</summary>
public sealed record CellBorder
{
    public BorderSide? Left { get; init; }
    public BorderSide? Right { get; init; }
    public BorderSide? Top { get; init; }
    public BorderSide? Bottom { get; init; }

    public static readonly CellBorder None = new();

    public CellBorder WithLeft(BorderSide side) => this with { Left = side };
    public CellBorder WithRight(BorderSide side) => this with { Right = side };
    public CellBorder WithTop(BorderSide side) => this with { Top = side };
    public CellBorder WithBottom(BorderSide side) => this with { Bottom = side };

    /// <summary>Sets all four edges to the same style/color in one call.</summary>
    public CellBorder WithAllSides(BorderSide side) => this with { Left = side, Right = side, Top = side, Bottom = side };
}

/// <summary>A small set of named formats covering the vast majority of real spreadsheets -
/// mirrors the F# core's <c>NumberFormat</c> named cases (its <c>Custom</c> raw-format-code
/// escape hatch is exposed separately via <see cref="CellStyle.WithCustomNumberFormat"/>
/// rather than as a case here, since a raw format string isn't really a "kind").</summary>
public enum NumberFormatKind
{
    General,
    Integer,
    TwoDecimal,
    Percentage,
    Currency,
    ShortDate,
    DateAndTime
}

/// <summary>
/// Cell formatting - font, fill, border, alignment, and number format. Immutable: every
/// <c>With*</c>/<c>As*</c> method returns a new <see cref="CellStyle"/> via a record
/// <c>with</c>-expression rather than mutating in place, so a style can be built up once
/// and safely reused/branched across many cells without aliasing surprises. Mirrors the F#
/// core's <c>CellStyle</c> record.
/// </summary>
public sealed record CellStyle
{
    public string? FontName { get; init; }

    /// <summary>Points.</summary>
    public double? FontSize { get; init; }

    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strikethrough { get; init; }
    public RgbColor? FontColor { get; init; }

    /// <summary>Solid fill color. Pattern fills (stripes, gradients) aren't modeled - see
    /// the F# core's own <c>FillStyle</c> doc comment.</summary>
    public RgbColor? FillColor { get; init; }

    public CellBorder? Border { get; init; }

    public HorizontalCellAlignment? HorizontalAlignment { get; init; }
    public VerticalCellAlignment? VerticalAlignment { get; init; }
    public bool WrapText { get; init; }

    public NumberFormatKind? NumberFormat { get; init; }

    /// <summary>A raw OOXML number format code (e.g. <c>"0.00%"</c>), used instead of <see
    /// cref="NumberFormat"/> when set. Setting one clears the other.</summary>
    public string? CustomNumberFormat { get; init; }

    /// <summary>Per-cell lock/hide flags - see <see cref="CellProtection"/>. <see
    /// langword="null"/> means Excel's own implicit default (locked, not hidden), same as an
    /// explicit <see cref="CellProtection.Default"/>.</summary>
    public CellProtection? Protection { get; init; }

    public static readonly CellStyle Default = new();

    public CellStyle WithFontName(string name) => this with { FontName = name };
    public CellStyle WithFontSize(double points) => this with { FontSize = points };
    public CellStyle AsBold() => this with { Bold = true };
    public CellStyle AsItalic() => this with { Italic = true };
    public CellStyle AsUnderline() => this with { Underline = true };
    public CellStyle AsStrikethrough() => this with { Strikethrough = true };
    public CellStyle WithFontColor(RgbColor color) => this with { FontColor = color };
    public CellStyle WithFillColor(RgbColor color) => this with { FillColor = color };
    public CellStyle WithBorder(CellBorder border) => this with { Border = border };
    public CellStyle WithHorizontalAlignment(HorizontalCellAlignment alignment) => this with { HorizontalAlignment = alignment };
    public CellStyle WithVerticalAlignment(VerticalCellAlignment alignment) => this with { VerticalAlignment = alignment };
    public CellStyle AsWrapText() => this with { WrapText = true };
    public CellStyle WithNumberFormat(NumberFormatKind format) => this with { NumberFormat = format, CustomNumberFormat = null };
    public CellStyle WithCustomNumberFormat(string formatCode) => this with { NumberFormat = null, CustomNumberFormat = formatCode };
    public CellStyle WithProtection(CellProtection protection) => this with { Protection = protection };
}
