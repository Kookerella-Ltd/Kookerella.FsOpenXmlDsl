namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A workbook-level named range/formula/constant - e.g. <c>"SalesData"</c> for a range, or a
/// named constant like <c>"TaxRate"</c> = <c>"0.075"</c>. <see cref="Formula"/> is raw
/// reference/formula text, the same convention as <see cref="CellValue.Formula"/>: whatever
/// Excel would show after the <c>=</c> - or, for a plain range reference (the common case),
/// no <c>=</c> involved at all, just the reference text itself, e.g.
/// <c>"Sheet1!$A$1:$B$10"</c>. Mirrors the F# core's <c>DefinedNameEntry</c>.
/// </summary>
public sealed record DefinedNameEntry(string Name, string Formula, DefinedNameScope Scope, bool Hidden = false)
{
    /// <summary>A workbook-scoped defined name, usable from any sheet.</summary>
    public static DefinedNameEntry Of(string name, string formula) => new(name, formula, new DefinedNameScope.WorkbookScope());

    /// <summary>A defined name restricted to one sheet.</summary>
    public static DefinedNameEntry Of(string name, string formula, string sheetName) => new(name, formula, new DefinedNameScope.SheetScope(sheetName));
}
