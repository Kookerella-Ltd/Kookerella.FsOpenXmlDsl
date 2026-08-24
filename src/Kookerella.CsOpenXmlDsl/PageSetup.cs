namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Print settings for a worksheet - orientation, paper size, scaling, margins, print area,
/// and a header/footer (including its first-page/even-page variants). Mirrors the F# core's
/// <c>PageSetup</c>. Immutable - every <c>With*</c> method returns a new instance.
/// <para>
/// <see cref="Header"/>/<see cref="Footer"/> (and their <c>Even</c>/<c>First</c>
/// counterparts) are raw OOXML header/footer text: Excel's own <c>&amp;L</c>/<c>&amp;C</c>/
/// <c>&amp;R</c> (left/center/right section) and <c>&amp;P</c>/<c>&amp;N</c>/<c>&amp;D</c>/
/// <c>&amp;T</c>/<c>&amp;F</c>/<c>&amp;A</c> (page number/total pages/date/time/filename/
/// sheet name) codes embedded directly in one string. <see cref="Header"/>/<see
/// cref="Footer"/> are shown on every page unless overridden: <see cref="EvenHeader"/>/<see
/// cref="EvenFooter"/> apply to even pages when set (Excel falls back to <see
/// cref="Header"/>/<see cref="Footer"/> for even pages otherwise), and <see
/// cref="FirstHeader"/>/<see cref="FirstFooter"/> apply only to page 1.
/// </para>
/// <para>
/// <see cref="PrintArea"/> is a list of ranges (Excel supports printing several disjoint
/// rectangles as one print area) - empty means "no print area set", i.e. Excel prints the
/// whole used range.
/// </para>
/// </summary>
public sealed record PageSetup
{
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;
    public PaperSize? PaperSize { get; init; }
    public PrintScaling? Scaling { get; init; }
    public PageMargins Margins { get; init; } = PageMargins.Default;
    public IReadOnlyList<(CellPosition TopLeft, CellPosition BottomRight)> PrintArea { get; init; } = Array.Empty<(CellPosition, CellPosition)>();
    public string? Header { get; init; }
    public string? Footer { get; init; }
    public string? EvenHeader { get; init; }
    public string? EvenFooter { get; init; }
    public string? FirstHeader { get; init; }
    public string? FirstFooter { get; init; }

    public static readonly PageSetup Default = new();

    public PageSetup WithOrientation(PageOrientation orientation) => this with { Orientation = orientation };
    public PageSetup WithPaperSize(PaperSize paperSize) => this with { PaperSize = paperSize };
    public PageSetup WithScaling(PrintScaling scaling) => this with { Scaling = scaling };
    public PageSetup WithMargins(PageMargins margins) => this with { Margins = margins };

    public PageSetup WithPrintArea(params (string TopLeftA1, string BottomRightA1)[] ranges) =>
        this with { PrintArea = ranges.Select(r => (CellPosition.FromA1(r.TopLeftA1), CellPosition.FromA1(r.BottomRightA1))).ToArray() };

    public PageSetup WithHeader(string header) => this with { Header = header };
    public PageSetup WithFooter(string footer) => this with { Footer = footer };
    public PageSetup WithEvenHeader(string header) => this with { EvenHeader = header };
    public PageSetup WithEvenFooter(string footer) => this with { EvenFooter = footer };
    public PageSetup WithFirstHeader(string header) => this with { FirstHeader = header };
    public PageSetup WithFirstFooter(string footer) => this with { FirstFooter = footer };
}
