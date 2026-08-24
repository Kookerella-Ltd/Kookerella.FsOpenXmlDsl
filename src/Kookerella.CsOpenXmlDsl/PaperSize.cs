namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A small named set of common paper sizes (OOXML's <c>ST_PaperSize</c> numeric codes) plus
/// <see cref="OtherPaperSize"/> as an escape hatch for any of the several dozen remaining
/// codes that aren't worth naming individually - a closed set of immutable cases, mirroring
/// the F# core's own <c>PaperSize</c> discriminated union (same "sealed hierarchy with a
/// private base constructor" pattern <see cref="CellValue"/> uses).
/// </summary>
public abstract record PaperSize
{
    private PaperSize() { }

    public sealed record Letter : PaperSize;

    public sealed record Legal : PaperSize;

    public sealed record Tabloid : PaperSize;

    public sealed record A3 : PaperSize;

    public sealed record A4 : PaperSize;

    public sealed record OtherPaperSize(int Code) : PaperSize;
}
