namespace SafeOpenXml

open System

[<AutoOpen>]
module Model =

    /// A cell's content. Note there is no `String of string` case that means anything other
    /// than plain text - rich text runs (multiple fonts/colors within a single cell's text)
    /// are a documented gap; a `Text` cell always renders/parses as one uniformly-styled run.
    type CellValue =
        | Empty
        | Text of string
        | Number of float
        | Boolean of bool
        | Date of DateTime
        /// `expression` excludes the leading "=". `cachedValue` is what Excel shows before
        /// it recalculates on open; leave it `None` to force a recalculation on first open.
        | Formula of expression: string * cachedValue: float option

    type Cell =
        { Ref: CellRef
          Value: CellValue
          Style: CellStyle option }

    /// Width is in OOXML's own column-width units (character widths of the workbook's
    /// default font) - there is no cleaner unit to convert to/from, so this is already 1:1.
    type ColumnProps = { Width: float option }

    /// Height is in points, matching OOXML directly.
    type RowProps = { Height: float option }

    type MergedRange =
        { TopLeft: CellRef
          BottomRight: CellRef }

    /// Number of leading rows/columns frozen in place, e.g. `{ Rows = 1; Columns = 0 }`
    /// freezes a header row.
    type FreezePane = { Rows: int; Columns: int }

    /// The range showing filter dropdown arrows. Core only models that the arrows are
    /// shown over this range - not any active filter criteria a user (or a foreign file)
    /// may have configured on top of them; see MAPPING.md.
    type AutoFilterRange =
        { TopLeft: CellRef
          BottomRight: CellRef }

    /// A worksheet's contents are a sparse set of addressed `Cell`s rather than nested
    /// rows-of-cells: real spreadsheets are sparse, and this avoids forcing every row to
    /// enumerate empty cells. Row/column metadata (height, width) is tracked separately
    /// since it can exist independent of any cell content in that row/column.
    type Worksheet =
        { Name: string
          Cells: Cell list
          ColumnProps: Map<int, ColumnProps>
          RowProps: Map<int, RowProps>
          MergedRanges: MergedRange list
          FreezePane: FreezePane option
          AutoFilter: AutoFilterRange option
          Protection: SheetProtection option
          ConditionalFormats: ConditionalFormatEntry list
          DataValidations: DataValidationEntry list
          Hyperlinks: HyperlinkEntry list
          Comments: CommentEntry list
          PageSetup: PageSetup option
          Tables: TableEntry list
          SparklineGroups: SparklineGroupEntry list
          Charts: ChartEntry list }

    type Workbook =
        { Sheets: Worksheet list
          DefinedNames: DefinedNameEntry list
          Protection: WorkbookProtection option }
