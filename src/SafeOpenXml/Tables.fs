namespace SafeOpenXml

/// One column of an Excel Table (`tableColumn`). `Name` must be unique within the table
/// and, per Excel's own requirement, should match the text of the header cell at that
/// column's position in the worksheet - Core doesn't synthesize that header cell for you,
/// the same way conditional formatting/autofilter/merges don't synthesize cell content
/// either, it only describes the range and metadata layered on top of cells you've
/// already placed.
type TableColumn =
    { Name: string
      /// A formula (same raw-text convention as `CellValue.Formula`) auto-filled down
      /// every data row of this column - OOXML's `calculatedColumnFormula`.
      CalculatedFormula: string option }

/// Visual style reference (`tableStyleInfo`) - a named built-in or custom table style plus
/// the banded-rows/-columns and first/last-column-emphasis toggles Excel's Table Design
/// ribbon exposes. `Name` is a raw style name (e.g. `"TableStyleMedium2"`, one of Excel's
/// several dozen built-ins, or a custom style name defined elsewhere in the workbook) -
/// Core doesn't model style *definitions* (their colors/fonts), only the reference to one.
type TableStyle =
    { Name: string option
      ShowFirstColumn: bool
      ShowLastColumn: bool
      ShowRowStripes: bool
      ShowColumnStripes: bool }

    /// Matches what Excel's own "Insert > Table" applies by default.
    static member Default =
        { Name = Some "TableStyleMedium2"
          ShowFirstColumn = false
          ShowLastColumn = false
          ShowRowStripes = true
          ShowColumnStripes = false }

/// An Excel Table (a `ListObject`, the thing structured references like `Table1[Column]`
/// point at) over a range, as stored on `Worksheet`. `Columns.Length` must equal the
/// range's width - one column per range column, left to right - and every `Name` must be
/// unique within the table (both are genuine Excel/OOXML requirements, not a
/// Core-specific restriction); `Writer` raises if either is violated rather than quietly
/// producing a file Excel would refuse to open cleanly.
///
/// Always has exactly a header row (`headerRowCount="1"`, OOXML's default and the
/// overwhelmingly common case) and never a totals row - see MAPPING.md for both gaps, and
/// for how `Name` doubles as OOXML's separate `name`/`displayName` attributes.
type TableEntry =
    { TopLeft: CellRef
      BottomRight: CellRef
      Name: string
      Columns: TableColumn list
      Style: TableStyle }
