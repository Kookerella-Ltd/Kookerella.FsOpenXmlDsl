namespace Kookerella.CsOpenXmlDsl;

/// <summary>The range showing filter dropdown arrows - mirrors the F# core's
/// <c>AutoFilterRange</c>. Only that the arrows are shown is modeled, not any active
/// filter criteria a user may configure on top of them afterward.</summary>
public sealed record AutoFilterRange(CellPosition TopLeft, CellPosition BottomRight)
{
    public static AutoFilterRange Of(string topLeftA1, string bottomRightA1) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1));
}
