namespace Kookerella.CsOpenXmlDsl;

/// <summary>Per-cell lock/hide flags. These only actually do anything once sheet-level
/// protection (<see cref="SheetProtection"/>) is turned on for the worksheet - Excel
/// ignores them otherwise. Every cell is <see cref="Locked"/> = <see langword="true"/> by
/// default even without an explicit <see cref="CellStyle"/> at all, matching Excel's own
/// default; <see cref="Hidden"/> (true = the cell's formula is hidden from the formula bar
/// once protected) defaults to <see langword="false"/>.</summary>
public sealed record CellProtection
{
    public bool Locked { get; init; } = true;
    public bool Hidden { get; init; }

    public static readonly CellProtection Default = new();

    public CellProtection WithLocked(bool locked = true) => this with { Locked = locked };
    public CellProtection WithHidden(bool hidden = true) => this with { Hidden = hidden };
}
