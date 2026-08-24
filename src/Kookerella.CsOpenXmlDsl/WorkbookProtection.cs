namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Workbook-level protection - protects the workbook's *structure* (sheet order,
/// visibility, adding/removing sheets) and/or window layout, as distinct from per-sheet
/// <see cref="SheetProtection"/> (which protects one sheet's own cell editing). Mirrors the
/// F# core's <c>WorkbookProtection</c>. Simpler than <see cref="SheetProtection"/>: both
/// flags here are plain "true means protected", no inverted-meaning trap to guard against.
/// Immutable - every <c>With*</c> method returns a new instance.
/// </summary>
public sealed record WorkbookProtection
{
    /// <summary>Plaintext password. Hashed with the same legacy XOR algorithm as
    /// <see cref="SheetProtection.Password"/>, for the same reasons (broadest compatibility,
    /// not real security) - never round-trips back to plaintext.</summary>
    public string? Password { get; init; }

    /// <summary>Prevents adding, deleting, hiding, unhiding, renaming, or reordering
    /// sheets.</summary>
    public bool? LockStructure { get; init; }

    /// <summary>Prevents moving or resizing the workbook's window - a legacy setting from
    /// when each workbook had its own window; rarely meaningful in modern Excel.</summary>
    public bool? LockWindows { get; init; }

    public static readonly WorkbookProtection Default = new();

    public WorkbookProtection WithPassword(string password) => this with { Password = password };
    public WorkbookProtection WithLockStructure(bool locked = true) => this with { LockStructure = locked };
    public WorkbookProtection WithLockWindows(bool locked = true) => this with { LockWindows = locked };
}
