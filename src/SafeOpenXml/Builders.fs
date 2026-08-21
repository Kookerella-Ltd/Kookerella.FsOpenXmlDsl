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
          FreezePane = None }

    /// Builds a `Worksheet` directly from a flat, pre-addressed cell list - for when your
    /// cells don't naturally arrive grouped by row (e.g. already `CellRef`-addressed data).
    /// See `sheet` for the row-grouped, tree-shaped alternative.
    let sheetOfCells (name: string) (cells: Cell list) : Worksheet =
        { emptySheet name with Cells = cells }

    let workbook (sheets: Worksheet list) : Workbook = { Sheets = sheets }

    let cellA1 (a1: string) value : Cell =
        { Ref = CellRef.ofA1 a1; Value = value; Style = None }

    let styledCellA1 (a1: string) value style : Cell =
        { Ref = CellRef.ofA1 a1; Value = value; Style = Some style }

    let withStyle style (c: Cell) : Cell = { c with Style = Some style }

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
type SheetItem =
    | Row of index: int option * cells: CellEntry list
    | ColumnWidth of index: int * width: float
    | RowHeight of index: int * height: float
    | Merge of topLeft: CellRef * bottomRight: CellRef
    | Freeze of rows: int * columns: int

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
          FreezePane = freezePaneOf items }
