namespace Kookerella.FsOpenXmlDsl

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
          Charts: ChartEntry list
          Images: ImageEntry list
          PivotTables: PivotTableEntry list }

    /// A VBA project (macros), stored as its own file's raw bytes exactly as they'd sit in
    /// `xl/vbaProject.bin` - a compiled OLE/CFBF binary blob, not source text. This DSL
    /// does no parsing, decompilation, or generation of VBA code (that would mean
    /// implementing a VBA compiler and the MS-OVBA binary format from scratch); it only
    /// embeds and hands back exactly the bytes you give it, the same "opaque payload"
    /// treatment `ImageEntry.Data` gets for raster images. Presence of a VBA project also
    /// switches the saved file's content type to Excel's macro-enabled kind (`Writer`
    /// picks `SpreadsheetDocumentType.MacroEnabledWorkbook`) - real Excel refuses to trust
    /// or run macros from a file whose content type doesn't declare them, regardless of the
    /// file's on-disk extension, so callers should also give the output path an `.xlsm`
    /// extension. See MAPPING.md for what this doesn't cover (authoring/editing macro
    /// source, digitally signed projects, wiring a macro to a form control or ribbon
    /// button).
    type Workbook =
        { Sheets: Worksheet list
          DefinedNames: DefinedNameEntry list
          Protection: WorkbookProtection option
          VbaProject: byte[] option }
