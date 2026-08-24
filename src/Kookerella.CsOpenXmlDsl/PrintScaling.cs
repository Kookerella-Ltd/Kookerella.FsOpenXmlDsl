namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// How the printed sheet is scaled onto its pages - Excel's print dialog toggles between
/// these two mutually exclusive modes ("Adjust to" vs. "Fit to"), a closed set of immutable
/// cases mirroring the F# core's own <c>PrintScaling</c> discriminated union. <see
/// cref="FitToPage"/>'s <see cref="FitToPage.Width"/>/<see cref="FitToPage.Height"/> are
/// page counts; <c>0</c> means "as many as needed" in that dimension, matching Excel's own
/// convention (e.g. "fit to 1 page wide, any number tall" is <c>new FitToPage(1, 0)</c>).
/// </summary>
public abstract record PrintScaling
{
    private PrintScaling() { }

    public sealed record ScalePercent(int Percent) : PrintScaling;

    public sealed record FitToPage(int Width, int Height) : PrintScaling;
}
