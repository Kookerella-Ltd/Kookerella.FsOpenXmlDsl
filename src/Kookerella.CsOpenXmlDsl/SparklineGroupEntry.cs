namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A group of sparklines sharing one <see cref="SparklineStyle"/> (a sheet can have several
/// independently-styled groups - e.g. one <see cref="SparklineType.Line"/> group and one
/// <see cref="SparklineType.Column"/> group in different columns). Mirrors the F# core's
/// <c>SparklineGroupEntry</c>. Immutable - <see cref="WithStyle"/> returns a new instance.
/// </summary>
public sealed record SparklineGroupEntry
{
    public SparklineStyle Style { get; init; } = SparklineStyle.Default;
    public IReadOnlyList<SparklineCell> Sparklines { get; init; } = Array.Empty<SparklineCell>();

    public SparklineGroupEntry(params SparklineCell[] sparklines) => Sparklines = sparklines;

    public SparklineGroupEntry WithStyle(SparklineStyle style) => this with { Style = style };
}
