namespace Kookerella.CsOpenXmlDsl;

/// <summary>The image's file format - determines the OOXML content type it's registered
/// with. Mirrors the F# core's <c>ImageFormat</c>: covers the four formats every Excel
/// version has always supported natively (TIFF/SVG/EMF/WMF aren't modeled in either
/// layer).</summary>
public enum ImageFormat
{
    Png,
    Jpeg,
    Gif,
    Bmp
}
