namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One sparkline's target cell and the data range it summarizes. <see cref="DataTopLeft"/>/
/// <see cref="DataBottomRight"/> are typically a single row or column (the common case: one
/// sparkline per row of data), but any rectangular range is accepted, the same as everywhere
/// else a range appears in this wrapper. Mirrors the F# core's <c>SparklineCell</c>.
/// </summary>
public sealed record SparklineCell(CellPosition Cell, CellPosition DataTopLeft, CellPosition DataBottomRight)
{
    public static SparklineCell Of(string cellA1, string dataTopLeftA1, string dataBottomRightA1) =>
        new(CellPosition.FromA1(cellA1), CellPosition.FromA1(dataTopLeftA1), CellPosition.FromA1(dataBottomRightA1));
}
