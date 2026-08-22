namespace SafeOpenXml

/// Plain functions for constructing DSL values directly (pipe/expression style).
[<AutoOpen>]
module Builders =

    let cellRef row col = CellRef.create row col

    let emptySheet (name: string) : Worksheet =
        { Name = name
          Cells = []
          ColumnProps = Map.empty
          RowProps = Map.empty
          MergedRanges = []
          FreezePane = None
          AutoFilter = None
          Protection = None
          ConditionalFormats = []
          DataValidations = []
          Hyperlinks = []
          Comments = []
          PageSetup = None
          Tables = []
          SparklineGroups = [] }

    /// Builds a `Worksheet` directly from a flat, pre-addressed cell list - for when your
    /// cells don't naturally arrive grouped by row (e.g. already `CellRef`-addressed data).
    /// See `sheet` for the row-grouped, tree-shaped alternative.
    let sheetOfCells (name: string) (cells: Cell list) : Worksheet =
        { emptySheet name with Cells = cells }

    let workbook (sheets: Worksheet list) : Workbook = { Sheets = sheets; DefinedNames = [] }

    let cellA1 (a1: string) value : Cell =
        { Ref = CellRef.ofA1 a1; Value = value; Style = None }

    let styledCellA1 (a1: string) value style : Cell =
        { Ref = CellRef.ofA1 a1; Value = value; Style = Some style }

    let withStyle style (c: Cell) : Cell = { c with Style = Some style }

    /// A workbook-scoped defined name, usable from any sheet.
    let definedName (name: string) (formula: string) : DefinedNameEntry =
        { Name = name; Formula = formula; Scope = WorkbookScope; Hidden = false }

    /// A defined name scoped to one sheet - shadows a workbook-scoped name of the same
    /// name when used from that sheet.
    let sheetScopedDefinedName (sheetName: string) (name: string) (formula: string) : DefinedNameEntry =
        { Name = name; Formula = formula; Scope = SheetScope sheetName; Hidden = false }

    /// Attaches defined names to a workbook - pipe-friendly: `workbook [...] |> withDefinedNames [...]`.
    let withDefinedNames (names: DefinedNameEntry list) (wb: Workbook) : Workbook = { wb with DefinedNames = names }

/// A single cell placed within a `Row` - one simple shape, no separate "styled" or
/// "explicit position" case. `Col = None` means "the next column after the previous entry
/// in this row" (starting at 0); `Some c` jumps to column c explicitly, and sequential
/// numbering resumes right after it. `Style = None` means no explicit style.
///
/// Build these with `SheetDsl.cell` (below) rather than the case directly, so the common
/// unstyled/sequential entry doesn't need to spell out two `None`s.
type CellEntry = Cell of col: int option * value: CellValue * style: CellStyle option

/// One fact about a worksheet - construct these directly as a flat list, in any order,
/// and `sheet` interprets them into the canonical `Worksheet`. This is the DSL's own
/// "AST for building a sheet", mirroring how SpreadsheetML itself is a tree of rows
/// containing cells, plus sheet-level metadata (column widths, merges, frozen panes).
///
/// `Row`'s `Index = None` means "the next row after the previous `Row`" (starting at 0) -
/// the common case, matching how you'd naturally list a sheet's rows top-to-bottom;
/// `Some r` jumps to row r explicitly, for gaps, and sequential numbering resumes right
/// after it. Same relationship as `CellEntry.Col`. Build these with `SheetDsl.row` (below).
///
/// `DocumentFormat.OpenXml.Spreadsheet` also defines a type named `Row`, and `Writer`
/// constructs it via `Spreadsheet.Row(...)` - qualified, precisely so it isn't shadowed by
/// this case now that `Writer` also has `open SafeOpenXml` in scope. `Reader`'s
/// `Elements<Row>()` calls need no such qualification: a type-argument position only ever
/// resolves to an actual type, never to a union case (case names live in the value
/// namespace, not the type namespace) - so those already meant the OOXML type unambiguously.
/// `DocumentFormat.OpenXml.Spreadsheet` also defines types named `ConditionalFormatting`/
/// `DataValidation`/`Hyperlink`/`Comment`/`AutoFilter` (note: singular "ConditionalFormat"
/// here vs. the OOXML type's plural/gerund "ConditionalFormatting", so those two don't
/// actually collide by name; the `DataValidation`, `Hyperlink`, `Comment`, and
/// `AutoFilter` cases genuinely do collide with the OOXML types of the same name, and
/// `Writer` qualifies their construction as `Spreadsheet.DataValidation(...)`/
/// `Spreadsheet.Hyperlink(...)`/`Spreadsheet.Comment(...)`/`Spreadsheet.AutoFilter(...)`
/// for exactly that reason, same as `Spreadsheet.Row`/`Spreadsheet.Cell` elsewhere.
/// `PageSetup` collides twice over - both with the OOXML type of the same name and with
/// this DSL's own `PageSetup` record (the case's payload type, same trick as `CellEntry`'s
/// `Cell` case sharing a name with `Model.Cell`) - so `Writer`/`Reader` always write
/// `Spreadsheet.PageSetup`/`Spreadsheet.PageMargins` explicitly. `Table` collides with
/// `DocumentFormat.OpenXml.Spreadsheet.Table` the same way `Row`/`Cell` do, qualified as
/// `Spreadsheet.Table` for the same reason. `SparklineGroup` collides with
/// `DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup` - that whole namespace is
/// aliased as `X14` in `Writer`/`Reader` (its own several-type surface, rather than one
/// qualifier, made a short alias clearer than repeating `Office2010.Excel.` everywhere).
type SheetItem =
    | Row of index: int option * cells: CellEntry list
    | ColumnWidth of index: int * width: float
    | RowHeight of index: int * height: float
    | Merge of topLeft: CellRef * bottomRight: CellRef
    | Freeze of rows: int * columns: int
    | AutoFilter of topLeft: CellRef * bottomRight: CellRef
    | Protect of settings: SheetProtection
    | ConditionalFormat of topLeft: CellRef * bottomRight: CellRef * rule: ConditionalFormatRule
    | DataValidation of topLeft: CellRef * bottomRight: CellRef * kind: ValidationKind * alert: ValidationAlert
    | Hyperlink of topLeft: CellRef * bottomRight: CellRef * target: HyperlinkTarget * tooltip: string option * display: string option
    | Comment of cell: CellRef * author: string * text: string
    | PageSetup of settings: PageSetup
    | Table of entry: TableEntry
    | SparklineGroup of entry: SparklineGroupEntry

/// Smart constructors for `CellEntry`/`SheetItem`, as members with real optional
/// parameters (`?col`, `?style`, `?index`) rather than several separately-named functions
/// for each combination - plain `let` bindings can't have optional parameters in F#
/// (that's a member-only feature), so this needs a type. `open type SafeOpenXml.SheetDsl`
/// alongside `open SafeOpenXml` brings `cell`/`row` into scope unqualified, the same
/// way `open` does for a module.
type SheetDsl =
    /// `col` defaults to the next column after the previous entry in the row;
    /// `style` defaults to no explicit style.
    static member cell(value: CellValue, ?col: int, ?style: CellStyle) : CellEntry =
        Cell(col, value, style)

    /// `index` defaults to the next row after the previous row in the sheet.
    static member row(cells: CellEntry list, ?index: int) : SheetItem =
        Row(index, cells)

    /// Shows filter dropdown arrows over the range - no active filter criteria, just the
    /// arrows (matching the common "Insert > filter" case where the criteria are left for
    /// whoever opens the file to set interactively).
    static member autoFilter(topLeft: CellRef, bottomRight: CellRef) : SheetItem =
        AutoFilter(topLeft, bottomRight)

    static member conditionalFormat(topLeft: CellRef, bottomRight: CellRef, rule: ConditionalFormatRule) : SheetItem =
        ConditionalFormat(topLeft, bottomRight, rule)

    /// `allowBlank` defaults to `true`; `errorStyle` defaults to `Stop`. The remaining
    /// optional parameters are the input prompt / error alert shown to the user - omit
    /// them all for a plain validation rule with no custom messaging.
    static member dataValidation
        (
            topLeft: CellRef,
            bottomRight: CellRef,
            kind: ValidationKind,
            ?allowBlank: bool,
            ?errorStyle: ErrorAlertStyle,
            ?errorTitle: string,
            ?errorMessage: string,
            ?inputTitle: string,
            ?inputMessage: string
        ) : SheetItem =
        DataValidation(
            topLeft,
            bottomRight,
            kind,
            { AllowBlank = defaultArg allowBlank true
              ErrorStyle = defaultArg errorStyle Stop
              ErrorTitle = errorTitle
              ErrorMessage = errorMessage
              InputTitle = inputTitle
              InputMessage = inputMessage }
        )

    /// Hyperlink over a single cell.
    static member hyperlink(cell: CellRef, target: HyperlinkTarget, ?tooltip: string, ?display: string) : SheetItem =
        Hyperlink(cell, cell, target, tooltip, display)

    /// Hyperlink over a range - every cell in it shares the same target.
    static member hyperlink
        (
            topLeft: CellRef,
            bottomRight: CellRef,
            target: HyperlinkTarget,
            ?tooltip: string,
            ?display: string
        ) : SheetItem =
        Hyperlink(topLeft, bottomRight, target, tooltip, display)

    /// `author` defaults to an empty (unnamed) author, matching Excel's own behavior.
    static member comment(cell: CellRef, text: string, ?author: string) : SheetItem =
        Comment(cell, defaultArg author "", text)

[<AutoOpen>]
module SheetItems =

    /// Extracts the `Row`/cell facts from a `SheetItem` list into a flat cell list,
    /// threading the "next row" cursor that `Row`'s `index = None` defaults from (and,
    /// within each row, the "next column" cursor `Cell`'s `col = None` defaults from)
    /// through a fold instead of mutable counters. Row order in the result doesn't matter
    /// - `Writer` groups/sorts cells by row and column itself - so rows are gathered as a
    /// list of per-row chunks and flattened at the end rather than merged into a lookup
    /// keyed by row index.
    let private cellsOf (items: SheetItem list) : Cell list =
        let cellsForRow (index: int) (cellEntries: CellEntry list) : Cell list =
            cellEntries
            |> List.fold
                (fun (nextCol, acc) (Cell(colOpt, value, style)) ->
                    let col = defaultArg colOpt nextCol
                    let cell = { Ref = cellRef index col; Value = value; Style = style }
                    (col + 1, cell :: acc))
                (0, [])
            |> snd
            |> List.rev

        items
        |> List.fold
            (fun (nextRow, rowsAcc) item ->
                match item with
                | Row(indexOpt, cellEntries) ->
                    let index = defaultArg indexOpt nextRow
                    (index + 1, cellsForRow index cellEntries :: rowsAcc)
                | _ -> (nextRow, rowsAcc))
            (0, [])
        |> snd
        |> List.rev
        |> List.concat

    /// Extracts `ColumnWidth` facts, keyed by column index - `Map.ofList` keeps the last
    /// entry for any repeated key, so a later width for the same column overwrites an
    /// earlier one.
    let private columnWidthsOf (items: SheetItem list) : Map<int, ColumnProps> =
        items
        |> List.choose (function
            | ColumnWidth(index, width) -> Some(index, { Width = Some width })
            | _ -> None)
        |> Map.ofList

    /// Extracts `RowHeight` facts, keyed by row index - same last-wins rule as `columnWidthsOf`.
    let private rowHeightsOf (items: SheetItem list) : Map<int, RowProps> =
        items
        |> List.choose (function
            | RowHeight(index, height) -> Some(index, { Height = Some height })
            | _ -> None)
        |> Map.ofList

    /// Extracts `Merge` facts - order doesn't matter, duplicates are harmless.
    let private mergedRangesOf (items: SheetItem list) : MergedRange list =
        items
        |> List.choose (function
            | Merge(topLeft, bottomRight) -> Some { TopLeft = topLeft; BottomRight = bottomRight }
            | _ -> None)

    /// Extracts the (at most one) `Freeze` fact - a later entry overwrites an earlier one.
    let private freezePaneOf (items: SheetItem list) : FreezePane option =
        items
        |> List.choose (function
            | Freeze(rows, columns) -> Some { Rows = rows; Columns = columns }
            | _ -> None)
        |> List.tryLast

    /// Extracts the (at most one) `AutoFilter` fact - a later entry overwrites an earlier
    /// one, same rule as `freezePaneOf` (only one `autoFilter` element is allowed per sheet).
    let private autoFilterOf (items: SheetItem list) : AutoFilterRange option =
        items
        |> List.choose (function
            | AutoFilter(topLeft, bottomRight) -> Some { TopLeft = topLeft; BottomRight = bottomRight }
            | _ -> None)
        |> List.tryLast

    /// Extracts the (at most one) `Protect` fact - a later entry overwrites an earlier
    /// one, same rule as `freezePaneOf`/`autoFilterOf` (only one `sheetProtection` element
    /// is allowed per sheet).
    let private sheetProtectionOf (items: SheetItem list) : SheetProtection option =
        items
        |> List.choose (function
            | Protect settings -> Some settings
            | _ -> None)
        |> List.tryLast

    /// Extracts `ConditionalFormat` facts, in order - order matters here (it becomes rule
    /// priority when writing), unlike `Merge`/`DataValidation`.
    let private conditionalFormatsOf (items: SheetItem list) : ConditionalFormatEntry list =
        items
        |> List.choose (function
            | ConditionalFormat(topLeft, bottomRight, rule) -> Some { TopLeft = topLeft; BottomRight = bottomRight; Rule = rule }
            | _ -> None)

    /// Extracts `DataValidation` facts - order doesn't matter.
    let private dataValidationsOf (items: SheetItem list) : DataValidationEntry list =
        items
        |> List.choose (function
            | DataValidation(topLeft, bottomRight, kind, alert) ->
                Some { TopLeft = topLeft; BottomRight = bottomRight; Kind = kind; Alert = alert }
            | _ -> None)

    /// Extracts `Hyperlink` facts - order doesn't matter.
    let private hyperlinksOf (items: SheetItem list) : HyperlinkEntry list =
        items
        |> List.choose (function
            | Hyperlink(topLeft, bottomRight, target, tooltip, display) ->
                Some
                    { TopLeft = topLeft
                      BottomRight = bottomRight
                      Target = target
                      Tooltip = tooltip
                      Display = display }
            | _ -> None)

    /// Extracts `Comment` facts - order doesn't matter.
    let private commentsOf (items: SheetItem list) : CommentEntry list =
        items
        |> List.choose (function
            | Comment(cell, author, text) -> Some { Cell = cell; Author = author; Text = text }
            | _ -> None)

    /// Extracts the (at most one) `PageSetup` fact - a later entry overwrites an earlier
    /// one, same rule as `freezePaneOf`/`autoFilterOf`/`sheetProtectionOf` (only one
    /// `pageSetup`/`pageMargins`/`headerFooter` triple is allowed per sheet).
    let private pageSetupOf (items: SheetItem list) : PageSetup option =
        items
        |> List.choose (function
            | PageSetup settings -> Some settings
            | _ -> None)
        |> List.tryLast

    /// Extracts `Table` facts - order doesn't matter, and (unlike `Freeze`/`AutoFilter`/
    /// `Protect`/`PageSetup`) a sheet can genuinely have several tables at once.
    let private tablesOf (items: SheetItem list) : TableEntry list =
        items
        |> List.choose (function
            | Table entry -> Some entry
            | _ -> None)

    /// Extracts `SparklineGroup` facts - order doesn't matter, same as `Table` (a sheet
    /// can genuinely have several independently-styled sparkline groups at once).
    let private sparklineGroupsOf (items: SheetItem list) : SparklineGroupEntry list =
        items
        |> List.choose (function
            | SparklineGroup entry -> Some entry
            | _ -> None)

    /// Interprets a flat list of `SheetItem` facts into the canonical `Worksheet` record.
    /// Each concern above is a small pure function over the same `items` list - no shared
    /// mutable state - so adding a new kind of fact later means adding a new function and
    /// one more field here, not growing this function.
    let sheet (name: string) (items: SheetItem list) : Worksheet =
        { Name = name
          Cells = cellsOf items
          ColumnProps = columnWidthsOf items
          RowProps = rowHeightsOf items
          MergedRanges = mergedRangesOf items
          FreezePane = freezePaneOf items
          AutoFilter = autoFilterOf items
          Protection = sheetProtectionOf items
          ConditionalFormats = conditionalFormatsOf items
          DataValidations = dataValidationsOf items
          Hyperlinks = hyperlinksOf items
          Comments = commentsOf items
          PageSetup = pageSetupOf items
          Tables = tablesOf items
          SparklineGroups = sparklineGroupsOf items }
