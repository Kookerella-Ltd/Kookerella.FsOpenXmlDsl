namespace Kookerella.CsOpenXmlDsl;

/// <summary>A conditional formatting rule applied over a range (a sheet can have several).
/// Mirrors the F# core's <c>ConditionalFormatEntry</c>.</summary>
public sealed record ConditionalFormatEntry(CellPosition TopLeft, CellPosition BottomRight, ConditionalFormatRule Rule)
{
    public static ConditionalFormatEntry Of(string topLeftA1, string bottomRightA1, ConditionalFormatRule rule) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1), rule);
}
