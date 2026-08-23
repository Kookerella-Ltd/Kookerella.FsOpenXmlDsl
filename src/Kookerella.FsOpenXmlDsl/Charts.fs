namespace Kookerella.FsOpenXmlDsl

/// The chart's visual shape. Named with a `Chart` prefix (`ChartColumn` not `Column`) so
/// it doesn't collide with `SparklineType.Column`/`.Line` - both types live in this same
/// namespace and would otherwise silently shadow each other, the same trap `PaperSize`'s
/// `OtherPaperSize` already documents. Covers the three most common chart kinds; see
/// MAPPING.md for what isn't modeled (scatter/area/stock/radar/etc., 3-D variants,
/// stacked/percent-stacked grouping).
type ChartType =
    | ChartColumn
    | ChartBar
    | ChartLine
    | ChartPie

/// One data series: the range of values it plots, and a reference to the cell that names
/// it (matching how a real Excel chart's series name is normally the data's header cell,
/// live-updating if that cell's text changes - not a static copy of it). All series in a
/// chart share the same category range (`ChartEntry.Categories`), matching Excel's own
/// most common "one category column, N value columns" layout.
type ChartSeries =
    { Name: CellRef
      ValuesTopLeft: CellRef
      ValuesBottomRight: CellRef }

/// A chart anchored over a range of cells on a worksheet, as stored on `Worksheet` (a
/// sheet can have several). `Categories` is the range of axis labels (bar/line) or slice
/// labels (pie) shared by every series. `TopLeftAnchor`/`BottomRightAnchor` size and
/// position the chart by spanning exactly that range of cells (a "move and size with
/// cells" anchor, snapped to cell boundaries) - matching how merged ranges/tables/
/// autofilter are already addressed elsewhere in this DSL, rather than modeling pixel-
/// precise floating position. `Title` is plain literal text, unlike series names - see
/// MAPPING.md.
type ChartEntry =
    { Type: ChartType
      Title: string option
      CategoriesTopLeft: CellRef
      CategoriesBottomRight: CellRef
      Series: ChartSeries list
      ShowLegend: bool
      TopLeftAnchor: CellRef
      BottomRightAnchor: CellRef }
