namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One worksheet - a name, its rows, a handful of sheet-level facts (merged ranges, frozen
/// panes, an autofilter range), Excel Tables, and charts. Immutable - every
/// <c>With*</c>/<c>Add*</c> method returns a new <see cref="Sheet"/> rather than mutating
/// in place. v1 scope is cells/formulas/basic styling/merged ranges/freeze panes/autofilter/
/// tables/charts only; images, pivot tables, conditional formatting, and everything else
/// the F# core models are out of scope here - reference Kookerella.FsOpenXmlDsl directly
/// for those.
/// </summary>
public sealed record Sheet
{
    public string Name { get; }
    public IReadOnlyList<Row> Rows { get; init; } = Array.Empty<Row>();
    public IReadOnlyList<MergedRange> MergedRanges { get; init; } = Array.Empty<MergedRange>();
    public FreezePane? FreezePane { get; init; }
    public AutoFilterRange? AutoFilter { get; init; }
    public IReadOnlyList<TableEntry> Tables { get; init; } = Array.Empty<TableEntry>();
    public IReadOnlyList<ChartEntry> Charts { get; init; } = Array.Empty<ChartEntry>();

    public Sheet(string name) => Name = name;

    public static Sheet Create(string name, params Row[] rows) => new(name) { Rows = rows };

    public Sheet WithRows(params Row[] rows) => this with { Rows = rows };

    public Sheet AddRow(Row row) => this with { Rows = Rows.Append(row).ToArray() };

    public Sheet WithMergedRanges(params MergedRange[] ranges) => this with { MergedRanges = ranges };

    public Sheet AddMergedRange(MergedRange range) => this with { MergedRanges = MergedRanges.Append(range).ToArray() };

    public Sheet WithFreezePane(FreezePane freezePane) => this with { FreezePane = freezePane };

    /// <summary>Convenience overload for <c>WithFreezePane(new FreezePane(rows, columns))</c>.</summary>
    public Sheet WithFreezePane(int rows, int columns) => this with { FreezePane = new FreezePane(rows, columns) };

    public Sheet WithAutoFilter(AutoFilterRange range) => this with { AutoFilter = range };

    public Sheet WithTables(params TableEntry[] tables) => this with { Tables = tables };

    public Sheet AddTable(TableEntry table) => this with { Tables = Tables.Append(table).ToArray() };

    public Sheet WithCharts(params ChartEntry[] charts) => this with { Charts = charts };

    public Sheet AddChart(ChartEntry chart) => this with { Charts = Charts.Append(chart).ToArray() };
}
