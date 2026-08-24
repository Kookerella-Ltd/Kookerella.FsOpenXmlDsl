namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Sheet-level protection settings - thin and honest rather than clever, mirroring the F#
/// core's own <c>SheetProtection</c>. <see cref="Sheet"/> is the one always-explicit master
/// switch (protect worksheet contents/cell editing); every other flag is nullable and, when
/// <see langword="null"/>, is simply omitted from the OOXML output entirely, letting Excel's
/// own schema default apply, rather than this wrapper guessing or hardcoding what
/// "unspecified" should mean.
/// <para>
/// <b>Every flag below except <see cref="Sheet"/> is deliberately named with a "Blocked"
/// suffix, diverging from the F# core's own plain field names (<c>FormatCells</c>,
/// <c>Sort</c>, etc.) on purpose</b>: in the underlying OOXML/F# model, setting one of these
/// to <see langword="true"/> does not "enable" that action - it *blocks* it once the sheet
/// is protected (e.g. the F# core's own <c>FormatCells = Some true</c> means formatting
/// cells becomes *blocked*, matching the raw <c>formatCells</c> OOXML attribute directly).
/// Getting that backwards would silently produce a working, schema-valid, but
/// wrong-behavior file that no amount of schema validation would ever catch, so the
/// property/method names here say "Blocked" explicitly rather than leave the direction to
/// be guessed or misremembered.
/// </para>
/// Per-cell lock/hide protection (<c>CellStyle.Protection</c> in the F# core) isn't exposed
/// at this layer - reference <c>Kookerella.FsOpenXmlDsl</c> directly for that; without it,
/// protecting a sheet through this wrapper alone locks every cell (Excel's own default),
/// with no way to unlock specific input cells first. Immutable - every <c>With*</c> method
/// returns a new instance.
/// </summary>
public sealed record SheetProtection
{
    /// <summary>Plaintext password. Hashed with the legacy (weak, XOR-based) algorithm on
    /// write, for the broadest possible Excel-version compatibility - this has never been
    /// real security, just a casual-editing speed bump, and that's still all it is here.
    /// Always reads back as <see langword="null"/>: the hash isn't reversible, so a
    /// round-tripped protected file loses password enforcement unless you re-supply the
    /// password.</summary>
    public string? Password { get; init; }

    public bool Sheet { get; init; } = true;

    public bool? ObjectsBlocked { get; init; }
    public bool? ScenariosBlocked { get; init; }
    public bool? FormatCellsBlocked { get; init; }
    public bool? FormatColumnsBlocked { get; init; }
    public bool? FormatRowsBlocked { get; init; }
    public bool? InsertColumnsBlocked { get; init; }
    public bool? InsertRowsBlocked { get; init; }
    public bool? InsertHyperlinksBlocked { get; init; }
    public bool? DeleteColumnsBlocked { get; init; }
    public bool? DeleteRowsBlocked { get; init; }
    public bool? SelectLockedCellsBlocked { get; init; }
    public bool? SortBlocked { get; init; }
    public bool? AutoFilterBlocked { get; init; }
    public bool? PivotTablesBlocked { get; init; }
    public bool? SelectUnlockedCellsBlocked { get; init; }

    public static readonly SheetProtection Default = new();

    public SheetProtection WithPassword(string password) => this with { Password = password };
    public SheetProtection WithSheet(bool protect = true) => this with { Sheet = protect };
    public SheetProtection WithObjectsBlocked(bool blocked = true) => this with { ObjectsBlocked = blocked };
    public SheetProtection WithScenariosBlocked(bool blocked = true) => this with { ScenariosBlocked = blocked };
    public SheetProtection WithFormatCellsBlocked(bool blocked = true) => this with { FormatCellsBlocked = blocked };
    public SheetProtection WithFormatColumnsBlocked(bool blocked = true) => this with { FormatColumnsBlocked = blocked };
    public SheetProtection WithFormatRowsBlocked(bool blocked = true) => this with { FormatRowsBlocked = blocked };
    public SheetProtection WithInsertColumnsBlocked(bool blocked = true) => this with { InsertColumnsBlocked = blocked };
    public SheetProtection WithInsertRowsBlocked(bool blocked = true) => this with { InsertRowsBlocked = blocked };
    public SheetProtection WithInsertHyperlinksBlocked(bool blocked = true) => this with { InsertHyperlinksBlocked = blocked };
    public SheetProtection WithDeleteColumnsBlocked(bool blocked = true) => this with { DeleteColumnsBlocked = blocked };
    public SheetProtection WithDeleteRowsBlocked(bool blocked = true) => this with { DeleteRowsBlocked = blocked };
    public SheetProtection WithSelectLockedCellsBlocked(bool blocked = true) => this with { SelectLockedCellsBlocked = blocked };
    public SheetProtection WithSortBlocked(bool blocked = true) => this with { SortBlocked = blocked };
    public SheetProtection WithAutoFilterBlocked(bool blocked = true) => this with { AutoFilterBlocked = blocked };
    public SheetProtection WithPivotTablesBlocked(bool blocked = true) => this with { PivotTablesBlocked = blocked };
    public SheetProtection WithSelectUnlockedCellsBlocked(bool blocked = true) => this with { SelectUnlockedCellsBlocked = blocked };
}
