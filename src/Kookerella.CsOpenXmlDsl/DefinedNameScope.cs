namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Where a defined name is visible - a closed set of immutable cases, mirroring the F#
/// core's own <c>DefinedNameScope</c> discriminated union (same "sealed hierarchy with a
/// private base constructor" pattern <see cref="CellValue"/> uses). <see
/// cref="WorkbookScope"/> makes it usable from any sheet; <see cref="SheetScope"/>
/// restricts it to one sheet, referenced by name.
/// </summary>
public abstract record DefinedNameScope
{
    private DefinedNameScope() { }

    public sealed record WorkbookScope : DefinedNameScope;

    public sealed record SheetScope(string SheetName) : DefinedNameScope;
}
