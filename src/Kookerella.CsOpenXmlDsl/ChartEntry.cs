namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A chart anchored over a range of cells on a worksheet (a sheet can have several).
/// <see cref="CategoriesTopLeft"/>/<see cref="CategoriesBottomRight"/> is the range of axis
/// labels (bar/line) or slice labels (pie) shared by every series. <see
/// cref="TopLeftAnchor"/>/<see cref="BottomRightAnchor"/> size and position the chart by
/// spanning exactly that range of cells - a "move and size with cells" anchor, snapped to
/// cell boundaries, matching how merged ranges/tables/autofilter are already addressed in
/// this wrapper, rather than pixel-precise floating position. <see cref="Title"/> is plain
/// literal text, unlike series names. Mirrors the F# core's <c>ChartEntry</c>. Immutable -
/// <see cref="WithTitle"/>/<see cref="WithLegend"/> each return a new instance.
/// </summary>
public sealed record ChartEntry
{
    public ChartType Type { get; }
    public string? Title { get; init; }
    public CellPosition CategoriesTopLeft { get; }
    public CellPosition CategoriesBottomRight { get; }
    public IReadOnlyList<ChartSeries> Series { get; init; } = Array.Empty<ChartSeries>();
    public bool ShowLegend { get; init; }
    public CellPosition TopLeftAnchor { get; }
    public CellPosition BottomRightAnchor { get; }

    public ChartEntry(
        ChartType type,
        CellPosition categoriesTopLeft,
        CellPosition categoriesBottomRight,
        CellPosition topLeftAnchor,
        CellPosition bottomRightAnchor,
        params ChartSeries[] series)
    {
        Type = type;
        CategoriesTopLeft = categoriesTopLeft;
        CategoriesBottomRight = categoriesBottomRight;
        TopLeftAnchor = topLeftAnchor;
        BottomRightAnchor = bottomRightAnchor;
        Series = series;
    }

    public static ChartEntry Of(
        ChartType type,
        string categoriesTopLeftA1,
        string categoriesBottomRightA1,
        string topLeftAnchorA1,
        string bottomRightAnchorA1,
        params ChartSeries[] series) =>
        new(
            type,
            CellPosition.FromA1(categoriesTopLeftA1),
            CellPosition.FromA1(categoriesBottomRightA1),
            CellPosition.FromA1(topLeftAnchorA1),
            CellPosition.FromA1(bottomRightAnchorA1),
            series);

    public ChartEntry WithTitle(string title) => this with { Title = title };

    public ChartEntry WithLegend(bool show = true) => this with { ShowLegend = show };
}
