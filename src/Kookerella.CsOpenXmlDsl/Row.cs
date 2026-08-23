namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One row of cells within a <see cref="Sheet"/>. Immutable - <see cref="AtIndex"/> returns
/// a new <see cref="Row"/> rather than mutating in place. <see cref="Index"/> defaults to
/// <see langword="null"/>, meaning "the next row after the previous row in the sheet"
/// (starting at 0), the same convention <see cref="Cell.Column"/> uses within a row - set
/// it explicitly via <see cref="AtIndex"/> to jump to a specific row, after which sequential
/// numbering resumes right after it.
/// </summary>
public sealed record Row
{
    public IReadOnlyList<Cell> Cells { get; init; } = Array.Empty<Cell>();
    public int? Index { get; init; }

    public static Row Of(params Cell[] cells) => new() { Cells = cells };

    public Row AtIndex(int index) => this with { Index = index };
}
