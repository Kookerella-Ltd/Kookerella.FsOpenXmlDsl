namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A plain RGB color. Immutable value type - the wrapper only ever authors RGB colors
/// (matching the F# core's own <c>Color.Rgb</c> case); reading back a foreign file's
/// theme/indexed colors isn't exposed at this layer - reference Kookerella.FsOpenXmlDsl
/// directly if you need that.
/// </summary>
public readonly record struct RgbColor(byte R, byte G, byte B)
{
    public static readonly RgbColor Black = new(0, 0, 0);
    public static readonly RgbColor White = new(255, 255, 255);
    public static readonly RgbColor Red = new(255, 0, 0);
    public static readonly RgbColor Green = new(0, 128, 0);
    public static readonly RgbColor Blue = new(0, 0, 255);
    public static readonly RgbColor Yellow = new(255, 255, 0);
}
