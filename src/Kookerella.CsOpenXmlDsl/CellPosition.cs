using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A zero-based cell address - <c>Row 0 / Column 0</c> is spreadsheet cell "A1", matching
/// the F# core's own <c>CellRef</c> exactly. Used for the range facts (<see
/// cref="MergedRange"/>, <see cref="AutoFilterRange"/>) rather than a raw <c>(int, int)</c>
/// tuple, so a call site reads as "top-left cell" rather than an anonymous pair of numbers.
/// A1-string conversion delegates to the F# core rather than re-implementing it.
/// </summary>
public readonly record struct CellPosition(int Row, int Column)
{
    public static CellPosition FromA1(string a1)
    {
        var cellRef = Fs.CellRefModule.ofA1(a1);
        return new CellPosition(cellRef.Row, cellRef.Col);
    }

    public string ToA1() => Fs.CellRefModule.toA1(new Fs.CellRef(Row, Column));
}
