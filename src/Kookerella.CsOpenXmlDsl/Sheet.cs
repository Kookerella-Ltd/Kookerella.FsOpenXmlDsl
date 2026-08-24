namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One worksheet - a name, its rows, a handful of sheet-level facts (merged ranges, frozen
/// panes, an autofilter range), Excel Tables, charts, images, pivot tables, sparkline
/// groups, conditional formatting rules, data validation rules, hyperlinks, and comments.
/// Immutable - every <c>With*</c>/<c>Add*</c> method returns a new <see cref="Sheet"/>
/// rather than mutating in place. v1 scope is cells/formulas/basic styling/merged ranges/
/// freeze panes/autofilter/tables/charts/images/pivot tables/sparklines/conditional
/// formatting/data validation/hyperlinks/comments only; print settings and everything else
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
    public IReadOnlyList<ImageEntry> Images { get; init; } = Array.Empty<ImageEntry>();
    public IReadOnlyList<PivotTableEntry> PivotTables { get; init; } = Array.Empty<PivotTableEntry>();
    public IReadOnlyList<SparklineGroupEntry> SparklineGroups { get; init; } = Array.Empty<SparklineGroupEntry>();
    public IReadOnlyList<ConditionalFormatEntry> ConditionalFormats { get; init; } = Array.Empty<ConditionalFormatEntry>();
    public IReadOnlyList<DataValidationEntry> DataValidations { get; init; } = Array.Empty<DataValidationEntry>();
    public IReadOnlyList<HyperlinkEntry> Hyperlinks { get; init; } = Array.Empty<HyperlinkEntry>();
    public IReadOnlyList<CommentEntry> Comments { get; init; } = Array.Empty<CommentEntry>();

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

    public Sheet WithImages(params ImageEntry[] images) => this with { Images = images };

    public Sheet AddImage(ImageEntry image) => this with { Images = Images.Append(image).ToArray() };

    public Sheet WithPivotTables(params PivotTableEntry[] pivotTables) => this with { PivotTables = pivotTables };

    public Sheet AddPivotTable(PivotTableEntry pivotTable) => this with { PivotTables = PivotTables.Append(pivotTable).ToArray() };

    public Sheet WithSparklineGroups(params SparklineGroupEntry[] sparklineGroups) => this with { SparklineGroups = sparklineGroups };

    public Sheet AddSparklineGroup(SparklineGroupEntry sparklineGroup) => this with { SparklineGroups = SparklineGroups.Append(sparklineGroup).ToArray() };

    public Sheet WithConditionalFormats(params ConditionalFormatEntry[] conditionalFormats) => this with { ConditionalFormats = conditionalFormats };

    public Sheet AddConditionalFormat(ConditionalFormatEntry conditionalFormat) => this with { ConditionalFormats = ConditionalFormats.Append(conditionalFormat).ToArray() };

    public Sheet WithDataValidations(params DataValidationEntry[] dataValidations) => this with { DataValidations = dataValidations };

    public Sheet AddDataValidation(DataValidationEntry dataValidation) => this with { DataValidations = DataValidations.Append(dataValidation).ToArray() };

    public Sheet WithHyperlinks(params HyperlinkEntry[] hyperlinks) => this with { Hyperlinks = hyperlinks };

    public Sheet AddHyperlink(HyperlinkEntry hyperlink) => this with { Hyperlinks = Hyperlinks.Append(hyperlink).ToArray() };

    public Sheet WithComments(params CommentEntry[] comments) => this with { Comments = comments };

    public Sheet AddComment(CommentEntry comment) => this with { Comments = Comments.Append(comment).ToArray() };
}
