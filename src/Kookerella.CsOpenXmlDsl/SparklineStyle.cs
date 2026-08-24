namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Shared visual settings for a group of sparklines - Excel groups sparklines so several can
/// be styled together at once (e.g. a filled-down column of row-summary sparklines). Covers
/// the commonly used subset: <see cref="Color"/> is the one color this wrapper models (the
/// main sparkline color itself; Excel's separate negative-point/axis/marker colors default to
/// its own automatic choices) and the <c>Show*</c> flags are the "highlight these points"
/// toggles from Excel's Sparkline Design ribbon. Mirrors the F# core's <c>SparklineStyle</c>.
/// Immutable - every <c>With*</c> method returns a new instance.
/// </summary>
public sealed record SparklineStyle
{
    public SparklineType Type { get; init; } = SparklineType.Line;
    public RgbColor? Color { get; init; }

    /// <summary>Points; only meaningful for <see cref="SparklineType.Line"/>.</summary>
    public double? LineWeight { get; init; }

    /// <summary>Only meaningful for <see cref="SparklineType.Line"/>.</summary>
    public bool ShowMarkers { get; init; }

    public bool ShowHigh { get; init; }
    public bool ShowLow { get; init; }
    public bool ShowFirst { get; init; }
    public bool ShowLast { get; init; }
    public bool ShowNegative { get; init; }

    public static readonly SparklineStyle Default = new();

    public SparklineStyle WithType(SparklineType type) => this with { Type = type };
    public SparklineStyle WithColor(RgbColor color) => this with { Color = color };
    public SparklineStyle WithLineWeight(double points) => this with { LineWeight = points };
    public SparklineStyle WithMarkers(bool show = true) => this with { ShowMarkers = show };
    public SparklineStyle WithHigh(bool show = true) => this with { ShowHigh = show };
    public SparklineStyle WithLow(bool show = true) => this with { ShowLow = show };
    public SparklineStyle WithFirst(bool show = true) => this with { ShowFirst = show };
    public SparklineStyle WithLast(bool show = true) => this with { ShowLast = show };
    public SparklineStyle WithNegative(bool show = true) => this with { ShowNegative = show };
}
