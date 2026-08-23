namespace Kookerella.FsOpenXmlDsl

/// Where a defined name is visible. `WorkbookScope` makes it usable from any sheet;
/// `SheetScope` restricts it to one sheet, referenced by name here - translated to
/// OOXML's `localSheetId` index at write time, and back again into a name on read.
type DefinedNameScope =
    | WorkbookScope
    | SheetScope of sheetName: string

/// A workbook-level named range/formula/constant - e.g. `"SalesData"` for a range, or a
/// named constant like `"TaxRate"` = `"0.075"`. `Formula` is raw reference/formula text,
/// the same convention as `CellValue.Formula` and everywhere else in this DSL: whatever
/// Excel would show after the `=` - or, for a plain range reference (the common case),
/// no `=` involved at all, just the reference text itself, e.g. `"Sheet1!$A$1:$B$10"`.
type DefinedNameEntry =
    { Name: string
      Formula: string
      Scope: DefinedNameScope
      Hidden: bool }
