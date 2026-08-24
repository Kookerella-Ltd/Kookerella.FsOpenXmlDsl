namespace Kookerella.CsOpenXmlDsl;

/// <summary>Margins in inches, matching OOXML's <c>pageMargins</c> directly - there is no
/// cleaner unit to convert to/from. Mirrors the F# core's <c>PageMargins</c>. Immutable -
/// every <c>With*</c> method returns a new instance.</summary>
public sealed record PageMargins
{
    public double Left { get; init; } = 0.7;
    public double Right { get; init; } = 0.7;
    public double Top { get; init; } = 0.75;
    public double Bottom { get; init; } = 0.75;
    public double Header { get; init; } = 0.3;
    public double Footer { get; init; } = 0.3;

    /// <summary>Excel's own built-in margins - what a fresh worksheet prints with even
    /// without explicit margins set at all.</summary>
    public static readonly PageMargins Default = new();

    public PageMargins WithLeft(double inches) => this with { Left = inches };
    public PageMargins WithRight(double inches) => this with { Right = inches };
    public PageMargins WithTop(double inches) => this with { Top = inches };
    public PageMargins WithBottom(double inches) => this with { Bottom = inches };
    public PageMargins WithHeader(double inches) => this with { Header = inches };
    public PageMargins WithFooter(double inches) => this with { Footer = inches };
}
