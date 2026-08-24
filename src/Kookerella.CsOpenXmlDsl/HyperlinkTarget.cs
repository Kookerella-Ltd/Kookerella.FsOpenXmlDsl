namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Where a hyperlink points - a closed set of immutable cases, mirroring the F# core's own
/// <c>HyperlinkTarget</c> discriminated union (same "sealed hierarchy with a private base
/// constructor" pattern <see cref="CellValue"/> uses). <see cref="ExternalHyperlink"/>
/// covers ordinary URLs and <c>mailto:</c> addresses alike - OOXML treats both as an
/// external relationship, just with a different URI scheme, so there's no separate email
/// case. <see cref="InternalHyperlink"/> is a same-workbook reference such as
/// <c>"Sheet2!A1"</c> or a defined name.
/// </summary>
public abstract record HyperlinkTarget
{
    private HyperlinkTarget() { }

    public sealed record ExternalHyperlink(string Url) : HyperlinkTarget;

    public sealed record InternalHyperlink(string Location) : HyperlinkTarget;
}
