namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One worksheet - a name plus its rows. Immutable - <see cref="WithRows"/> and <see
/// cref="AddRow"/> each return a new <see cref="Sheet"/> rather than mutating in place.
/// v1 scope is cells/formulas/basic styling only; tables, charts, images, pivot tables,
/// conditional formatting, and everything else the F# core models are out of scope here -
/// reference Kookerella.FsOpenXmlDsl directly for those.
/// </summary>
public sealed record Sheet
{
    public string Name { get; }
    public IReadOnlyList<Row> Rows { get; init; } = Array.Empty<Row>();

    public Sheet(string name) => Name = name;

    public static Sheet Create(string name, params Row[] rows) => new(name) { Rows = rows };

    public Sheet WithRows(params Row[] rows) => this with { Rows = rows };

    public Sheet AddRow(Row row) => this with { Rows = Rows.Append(row).ToArray() };
}
