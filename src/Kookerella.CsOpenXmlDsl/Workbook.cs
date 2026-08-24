namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A workbook - an ordered list of sheets, defined names, plus an optional VBA project. Pure
/// data, same as the F# core's own <c>Workbook</c> - deliberately has no <c>Save</c>/
/// <c>Load</c> methods on it, since those are I/O (see <see cref="WorkbookIO"/>, the one
/// place this wrapper does anything side-effecting). Immutable - <see cref="AddSheet"/>/<see
/// cref="WithDefinedNames"/>/<see cref="AddDefinedName"/>/<see cref="WithVbaProject"/> each
/// return a new <see cref="Workbook"/> rather than mutating in place.
/// </summary>
public sealed record Workbook
{
    private readonly byte[]? _vbaProject;

    public IReadOnlyList<Sheet> Sheets { get; init; } = Array.Empty<Sheet>();

    public IReadOnlyList<DefinedNameEntry> DefinedNames { get; init; } = Array.Empty<DefinedNameEntry>();

    /// <summary>
    /// A VBA project (macros) as the raw bytes of an <c>xl/vbaProject.bin</c> - a compiled
    /// OLE/CFBF binary, not source text. Nothing in this stack parses, generates, or edits
    /// VBA; the bytes are embedded and handed back verbatim. Set it via <see
    /// cref="WithVbaProject"/> and save to an <c>.xlsm</c> path (the file's content type
    /// switches to macro-enabled automatically, but real Excel expects the extension to
    /// match before it will trust and run macros).
    /// <para>
    /// This is the one property on an otherwise-pure type that can't fully guarantee
    /// immutability: <see cref="WithVbaProject"/> defensively copies what you pass in, so
    /// mutating your original array afterwards can't affect this workbook - but the array
    /// returned from this getter is the record's own, so callers must not write to it.
    /// </para>
    /// </summary>
    public byte[]? VbaProject
    {
        get => _vbaProject;
        init => _vbaProject = value;
    }

    public static Workbook Create(params Sheet[] sheets) => new() { Sheets = sheets };

    public Workbook AddSheet(Sheet sheet) => this with { Sheets = Sheets.Append(sheet).ToArray() };

    public Workbook WithDefinedNames(params DefinedNameEntry[] definedNames) => this with { DefinedNames = definedNames };

    public Workbook AddDefinedName(DefinedNameEntry definedName) => this with { DefinedNames = DefinedNames.Append(definedName).ToArray() };

    /// <summary>Attaches a VBA project, defensively copying <paramref name="vbaProjectBytes"/>
    /// so later mutations to the caller's array don't leak into this workbook.</summary>
    public Workbook WithVbaProject(byte[] vbaProjectBytes) =>
        this with { VbaProject = vbaProjectBytes.ToArray() };
}
