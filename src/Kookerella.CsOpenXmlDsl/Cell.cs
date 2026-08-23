namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A single cell within a <see cref="Row"/>. Immutable - <see cref="WithStyle"/> and <see
/// cref="AtColumn"/> each return a new <see cref="Cell"/> rather than mutating in place.
/// <see cref="Column"/> defaults to <see langword="null"/>, meaning "the next column after
/// the previous cell in this row" (starting at 0) - the same convention the F# core's
/// <c>CellEntry</c> uses; set it explicitly via <see cref="AtColumn"/> to jump to a specific
/// column, after which sequential numbering resumes right after it.
/// </summary>
public sealed record Cell
{
    public CellValue Value { get; }
    public CellStyle? Style { get; init; }
    public int? Column { get; init; }

    public Cell(CellValue value) => Value = value;

    public static Cell Text(string value) => new(new CellValue.Text(value));
    public static Cell Number(double value) => new(new CellValue.Number(value));
    public static Cell Boolean(bool value) => new(new CellValue.Boolean(value));
    public static Cell Date(DateTime value) => new(new CellValue.Date(value));

    public static Cell Formula(string expression, double? cachedValue = null) =>
        new(new CellValue.Formula(expression, cachedValue));

    public Cell WithStyle(CellStyle style) => this with { Style = style };
    public Cell AtColumn(int column) => this with { Column = column };
}
