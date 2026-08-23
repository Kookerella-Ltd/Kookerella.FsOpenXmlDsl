namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One data series: the range of values it plots, and a reference to the cell that names
/// it (matching how a real Excel chart's series name is normally the data's header cell,
/// live-updating if that cell's text changes - not a static copy). All series in a chart
/// share the same category range (<see cref="ChartEntry.CategoriesTopLeft"/>/<see
/// cref="ChartEntry.CategoriesBottomRight"/>). Mirrors the F# core's <c>ChartSeries</c>.
/// </summary>
public sealed record ChartSeries(CellPosition Name, CellPosition ValuesTopLeft, CellPosition ValuesBottomRight)
{
    public static ChartSeries Of(string nameA1, string valuesTopLeftA1, string valuesBottomRightA1) =>
        new(CellPosition.FromA1(nameA1), CellPosition.FromA1(valuesTopLeftA1), CellPosition.FromA1(valuesBottomRightA1));
}
