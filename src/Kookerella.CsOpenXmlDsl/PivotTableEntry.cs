namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A pivot table computed from a range of source cells and displayed as a grid anchored at
/// <see cref="TopLeftAnchor"/> (a sheet can have several). Unlike every other feature in this
/// wrapper, this one isn't a pure translation into OOXML - a pivot table's file format bakes
/// in the *result* of an aggregation in three places that all have to agree, so
/// <see cref="WorkbookIO.Save(Workbook, string)"/> actually performs the grouping and
/// aggregation described here rather than just describing a reference for Excel to resolve
/// later.
/// <para>
/// Scoped to the single most common pivot table shape, same as the F# core: exactly one row
/// field, at most one column field, and exactly one value field - no nested row/column
/// fields, no page/filter fields, no subtotals beyond the grand total row/column.
/// </para>
/// Mirrors the F# core's <c>PivotTableEntry</c>. Immutable - <see cref="WithSourceSheet"/>/
/// <see cref="WithColumnField"/>/<see cref="WithAggregation"/>/<see cref="WithValueCaption"/>
/// each return a new instance.
/// </summary>
public sealed record PivotTableEntry
{
    /// <summary><c>null</c> means the source range is on the same sheet as the pivot table
    /// itself.</summary>
    public string? SourceSheet { get; init; }

    public CellPosition SourceTopLeft { get; }
    public CellPosition SourceBottomRight { get; }

    /// <summary>Must exactly match a header cell's text in the source range's first row.</summary>
    public string RowField { get; }

    /// <summary>Must exactly match a header cell's text in the source range's first row, if set.</summary>
    public string? ColumnField { get; init; }

    /// <summary>Must exactly match a header cell's text in the source range's first row.</summary>
    public string ValueField { get; }

    public PivotAggregation Aggregation { get; init; } = PivotAggregation.Sum;

    /// <summary>Column header for the aggregated value - defaults to Excel's own convention
    /// (e.g. "Sum of Sales") when <c>null</c>.</summary>
    public string? ValueCaption { get; init; }

    public CellPosition TopLeftAnchor { get; }

    public PivotTableEntry(
        CellPosition sourceTopLeft,
        CellPosition sourceBottomRight,
        string rowField,
        string valueField,
        CellPosition topLeftAnchor)
    {
        SourceTopLeft = sourceTopLeft;
        SourceBottomRight = sourceBottomRight;
        RowField = rowField;
        ValueField = valueField;
        TopLeftAnchor = topLeftAnchor;
    }

    public static PivotTableEntry Of(
        string sourceTopLeftA1,
        string sourceBottomRightA1,
        string rowField,
        string valueField,
        string topLeftAnchorA1) =>
        new(
            CellPosition.FromA1(sourceTopLeftA1),
            CellPosition.FromA1(sourceBottomRightA1),
            rowField,
            valueField,
            CellPosition.FromA1(topLeftAnchorA1));

    public PivotTableEntry WithSourceSheet(string sheetName) => this with { SourceSheet = sheetName };

    public PivotTableEntry WithColumnField(string columnField) => this with { ColumnField = columnField };

    public PivotTableEntry WithAggregation(PivotAggregation aggregation) => this with { Aggregation = aggregation };

    public PivotTableEntry WithValueCaption(string caption) => this with { ValueCaption = caption };
}
