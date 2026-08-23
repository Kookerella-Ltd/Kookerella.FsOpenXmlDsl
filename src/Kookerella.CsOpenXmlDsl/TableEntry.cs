namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// An Excel Table (a <c>ListObject</c>, the thing structured references like
/// <c>Table1[Column]</c> point at) over a range. <see cref="Columns"/>' count must equal
/// the range's width - one column per range column, left to right - and every <see
/// cref="TableColumn.Name"/> must be unique within the table (both are genuine Excel/OOXML
/// requirements, not specific to this wrapper); saving raises if either is violated rather
/// than quietly producing a file Excel would refuse to open cleanly. Always has exactly a
/// header row and never a totals row. Mirrors the F# core's <c>TableEntry</c>. Immutable -
/// <see cref="WithStyle"/> returns a new instance.
/// </summary>
public sealed record TableEntry
{
    public CellPosition TopLeft { get; }
    public CellPosition BottomRight { get; }
    public string Name { get; }
    public IReadOnlyList<TableColumn> Columns { get; init; } = Array.Empty<TableColumn>();
    public TableStyle Style { get; init; } = TableStyle.Default;

    public TableEntry(CellPosition topLeft, CellPosition bottomRight, string name, params TableColumn[] columns)
    {
        TopLeft = topLeft;
        BottomRight = bottomRight;
        Name = name;
        Columns = columns;
    }

    public static TableEntry Of(string topLeftA1, string bottomRightA1, string name, params TableColumn[] columns) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1), name, columns);

    public TableEntry WithStyle(TableStyle style) => this with { Style = style };
}
