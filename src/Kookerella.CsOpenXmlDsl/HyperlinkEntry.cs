namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A hyperlink applied over a range (a single cell is the degenerate case where <see
/// cref="TopLeft"/> equals <see cref="BottomRight"/> - see <see cref="Of(string,
/// HyperlinkTarget)"/>). Mirrors the F# core's <c>HyperlinkEntry</c>. Immutable - <see
/// cref="WithTooltip"/>/<see cref="WithDisplay"/> each return a new instance.
/// </summary>
public sealed record HyperlinkEntry(CellPosition TopLeft, CellPosition BottomRight, HyperlinkTarget Target)
{
    /// <summary>OOXML's fallback label - a handful of older Excel versions and some
    /// interop tools show it instead of the cell's own text; modern Excel always shows the
    /// cell's actual content and ignores it. Rarely worth setting, but preserved for
    /// round-trip fidelity when reading a file that has one.</summary>
    public string? Display { get; init; }

    public string? Tooltip { get; init; }

    public static HyperlinkEntry Of(string topLeftA1, string bottomRightA1, HyperlinkTarget target) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1), target);

    /// <summary>Convenience overload for a hyperlink over a single cell.</summary>
    public static HyperlinkEntry Of(string cellA1, HyperlinkTarget target) => Of(cellA1, cellA1, target);

    public HyperlinkEntry WithTooltip(string tooltip) => this with { Tooltip = tooltip };

    public HyperlinkEntry WithDisplay(string display) => this with { Display = display };
}
