namespace Kookerella.CsOpenXmlDsl;

/// <summary>A rectangular range of cells merged into one - mirrors the F# core's
/// <c>MergedRange</c>.</summary>
public sealed record MergedRange(CellPosition TopLeft, CellPosition BottomRight)
{
    public static MergedRange Of(string topLeftA1, string bottomRightA1) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1));
}
