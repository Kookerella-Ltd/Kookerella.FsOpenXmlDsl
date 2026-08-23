namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A workbook - an ordered list of sheets. Pure data, same as the F# core's own
/// <c>Workbook</c> - deliberately has no <c>Save</c>/<c>Load</c> methods on it, since those
/// are I/O (see <see cref="WorkbookIO"/>, the one place this wrapper does anything
/// side-effecting). Immutable - <see cref="AddSheet"/> returns a new <see cref="Workbook"/>
/// rather than mutating in place.
/// </summary>
public sealed record Workbook
{
    public IReadOnlyList<Sheet> Sheets { get; init; } = Array.Empty<Sheet>();

    public static Workbook Create(params Sheet[] sheets) => new() { Sheets = sheets };

    public Workbook AddSheet(Sheet sheet) => this with { Sheets = Sheets.Append(sheet).ToArray() };
}
