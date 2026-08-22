namespace SafeOpenXml

/// Sheet-level protection settings - thin and honest rather than clever: `Sheet` is the
/// one always-explicit master switch (protect worksheet contents/cell editing against
/// the per-cell `CellProtection` locks); every other flag is `bool option` and, when
/// `None`, is simply omitted from the OOXML output entirely, letting Excel's own schema
/// default apply, rather than SafeOpenXml guessing or hardcoding what "unspecified"
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
