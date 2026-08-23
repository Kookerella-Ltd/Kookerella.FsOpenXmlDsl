namespace Kookerella.FsOpenXmlDsl

/// Sheet-level protection settings - thin and honest rather than clever: `Sheet` is the
/// one always-explicit master switch (protect worksheet contents/cell editing against
/// the per-cell `CellProtection` locks); every other flag is `bool option` and, when
/// `None`, is simply omitted from the OOXML output entirely, letting Excel's own schema
/// default apply, rather than Kookerella.FsOpenXmlDsl guessing or hardcoding what "unspecified"
/// should mean.
///
/// This is deliberate caution, not laziness: several of these flags are NOT "true enables
/// the action" - e.g. `FormatCells = Some true` means formatting cells is *blocked* once
/// protected, matching the underlying `formatCells` OOXML attribute directly. Getting a
/// default backwards here would silently produce a working, schema-valid, but
/// wrong-behavior file that no amount of schema validation would ever catch - so rather
/// than risk that, an unset flag is simply never written, and Excel's own documented
/// default takes over exactly as it would for a file Excel itself produced with that
/// attribute omitted.
type SheetProtection =
    { /// Plaintext password. Hashed with the legacy (weak, XOR-based) algorithm on write,
      /// for the broadest possible Excel-version compatibility - this has never been real
      /// security, just a casual-editing speed bump, and that's still all it is here.
      /// Always reads back as `None`: the hash isn't reversible, so a round-tripped
      /// protected file loses password enforcement unless you re-supply the password.
      Password: string option
      Sheet: bool
      Objects: bool option
      Scenarios: bool option
      FormatCells: bool option
      FormatColumns: bool option
      FormatRows: bool option
      InsertColumns: bool option
      InsertRows: bool option
      InsertHyperlinks: bool option
      DeleteColumns: bool option
      DeleteRows: bool option
      SelectLockedCells: bool option
      Sort: bool option
      AutoFilter: bool option
      PivotTables: bool option
      SelectUnlockedCells: bool option }

    static member Default =
        { Password = None
          Sheet = true
          Objects = None
          Scenarios = None
          FormatCells = None
          FormatColumns = None
          FormatRows = None
          InsertColumns = None
          InsertRows = None
          InsertHyperlinks = None
          DeleteColumns = None
          DeleteRows = None
          SelectLockedCells = None
          Sort = None
          AutoFilter = None
          PivotTables = None
          SelectUnlockedCells = None }

/// Workbook-level protection - protects the workbook's *structure* (sheet order,
/// visibility, adding/removing sheets) and/or window layout, as distinct from per-sheet
/// `SheetProtection` (which protects one sheet's own cell editing). Stored on `Workbook`
/// rather than `Worksheet` - the one other DSL concept, alongside `DefinedNameEntry`,
/// that's genuinely workbook-level.
///
/// Simpler than `SheetProtection`: both flags here are plain "true means protected", with
/// no inverted-meaning trap to guard against - but they're still `bool option`, only
/// written when the caller sets them explicitly, for the same reason as everywhere else
/// in this DSL: don't guess what "unspecified" should mean when Excel's own default
/// (unprotected) is right there and always safe.
type WorkbookProtection =
    { /// Plaintext password. Hashed with the same legacy XOR algorithm as
      /// `SheetProtection.Password`, for the same reasons (broadest compatibility, not
      /// real security) - never round-trips back to plaintext.
      Password: string option
      /// Prevents adding, deleting, hiding, unhiding, renaming, or reordering sheets.
      LockStructure: bool option
      /// Prevents moving or resizing the workbook's window - a legacy setting from when
      /// each workbook had its own window; rarely meaningful in modern Excel.
      LockWindows: bool option }

    static member Default =
        { Password = None
          LockStructure = None
          LockWindows = None }
