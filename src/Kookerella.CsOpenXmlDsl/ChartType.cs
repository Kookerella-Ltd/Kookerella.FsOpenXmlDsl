namespace Kookerella.CsOpenXmlDsl;

/// <summary>The chart's visual shape - mirrors the F# core's <c>ChartType</c>. Covers the
/// three most common chart kinds; scatter/area/stock/radar/etc., 3-D variants, and stacked/
/// percent-stacked grouping aren't modeled here or in the F# core.</summary>
public enum ChartType
{
    Column,
    Bar,
    Line,
    Pie
}
