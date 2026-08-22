namespace SafeOpenXml.Interpreter

open System
open System.Globalization
open System.Text
open SafeOpenXml

/// Renders a `Workbook` back out as an F# script that regenerates an equivalent file when
/// run - the reverse of `Reader` one level further: `Reader` turns OOXML into the DSL,
/// this turns the DSL into DSL *source text*. Every renderer below is a direct,
/// mechanical mirror of a DSL type's own shape (diffing against `.Default`/`.None` where
/// one exists, so generated code only mentions what isn't already implied) - there's no
/// separate "codegen model", just string-building over the same types `Builders`/`Model`
/// define.
module internal CodeGen =

    let private renderString (s: string) : string =
        s
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t")
        |> sprintf "\"%s\""

    /// Always includes a decimal point (or exponent) so F# infers `float`, not `int`.
    let private renderFloat (f: float) : string =
        let s = f.ToString("R", CultureInfo.InvariantCulture)
        if s.Contains "." || s.Contains "e" || s.Contains "E" then s else s + ".0"

    let private renderByte (b: byte) : string = sprintf "%duy" (int b)

    let private renderBool (b: bool) : string = if b then "true" else "false"

    /// Round-trips through the same `ToOADate`/`FromOADate` convention `Writer`/`Reader`
    /// already use for `CellValue.Date`, so generated code preserves exactly the precision
    /// the DSL itself preserves - no separate date-formatting logic to keep in sync.
    let private renderDateTime (d: DateTime) : string =
        sprintf "(System.DateTime.FromOADate(%s))" (renderFloat (d.ToOADate()))

    /// Always parenthesizes the wrapped value, not just the whole `Some (...)` - e.g.
    /// `Some Rgb(255uy, 255uy, 255uy)` is a parse error (`Rgb(...)`'s tupled-call syntax
    /// needs its own parens when passed as an argument), so this must render
    /// `Some (Rgb(255uy, 255uy, 255uy))`.
    let private renderOption (render: 'a -> string) (opt: 'a option) : string =
        match opt with
        | None -> "None"
        | Some v -> sprintf "(Some (%s))" (render v)

    let private renderCellRef (r: CellRef) : string =
        sprintf "(CellRef.ofA1 %s)" (renderString (CellRef.toA1 r))

    let private renderColor (c: Color) : string =
        match c with
        | Rgb(r, g, b) -> sprintf "Rgb(%s, %s, %s)" (renderByte r) (renderByte g) (renderByte b)
        | Indexed i -> sprintf "Indexed %d" i
        | Theme(i, tint) -> sprintf "Theme(%d, %s)" i (renderOption renderFloat tint)

    let private renderFontStyle (f: FontStyle) : string =
        if f = FontStyle.Default then
            "FontStyle.Default"
        else
            [ if f.Name.IsSome then yield sprintf "Name = %s" (renderOption renderString f.Name)
              if f.Size.IsSome then yield sprintf "Size = %s" (renderOption renderFloat f.Size)
              if f.Bold then yield "Bold = true"
              if f.Italic then yield "Italic = true"
              if f.Underline then yield "Underline = true"
              if f.Strikethrough then yield "Strikethrough = true"
              if f.Color.IsSome then yield sprintf "Color = %s" (renderOption renderColor f.Color) ]
            |> String.concat "; "
            |> sprintf "{ FontStyle.Default with %s }"

    let private renderFillStyle (f: FillStyle) : string = sprintf "{ Color = %s }" (renderColor f.Color)

    let private renderBorderLineStyle (s: BorderLineStyle) : string =
        match s with
        | Thin -> "Thin"
        | Medium -> "Medium"
        | Thick -> "Thick"
        | Dashed -> "Dashed"
        | Dotted -> "Dotted"
        | Double -> "Double"
        | Hair -> "Hair"
        | Other name -> sprintf "Other %s" (renderString name)

    let private renderBorderSide (s: BorderSide) : string =
        sprintf "{ Style = %s; Color = %s }" (renderBorderLineStyle s.Style) (renderOption renderColor s.Color)

    let private renderBorderStyle (b: BorderStyle) : string =
        if b = BorderStyle.None then
            "BorderStyle.None"
        else
            [ if b.Left.IsSome then yield sprintf "Left = %s" (renderOption renderBorderSide b.Left)
              if b.Right.IsSome then yield sprintf "Right = %s" (renderOption renderBorderSide b.Right)
              if b.Top.IsSome then yield sprintf "Top = %s" (renderOption renderBorderSide b.Top)
              if b.Bottom.IsSome then yield sprintf "Bottom = %s" (renderOption renderBorderSide b.Bottom) ]
            |> String.concat "; "
            |> sprintf "{ BorderStyle.None with %s }"

    let private renderHorizontalAlignment (a: HorizontalAlignment) : string =
        match a with
        | GeneralAlign -> "GeneralAlign"
        | AlignLeft -> "AlignLeft"
        | AlignCenter -> "AlignCenter"
        | AlignRight -> "AlignRight"
        | AlignFill -> "AlignFill"
        | AlignJustify -> "AlignJustify"

    let private renderVerticalAlignment (a: VerticalAlignment) : string =
        match a with
        | AlignTop -> "AlignTop"
        | AlignMiddle -> "AlignMiddle"
        | AlignBottom -> "AlignBottom"

    let private renderAlignmentStyle (a: AlignmentStyle) : string =
        if a = AlignmentStyle.Default then
            "AlignmentStyle.Default"
        else
            [ if a.Horizontal.IsSome then
                  yield sprintf "Horizontal = %s" (renderOption renderHorizontalAlignment a.Horizontal)
              if a.Vertical.IsSome then
                  yield sprintf "Vertical = %s" (renderOption renderVerticalAlignment a.Vertical)
              if a.WrapText then yield "WrapText = true" ]
            |> String.concat "; "
            |> sprintf "{ AlignmentStyle.Default with %s }"

    let private renderNumberFormat (n: NumberFormat) : string =
        match n with
        | General -> "General"
        | Integer -> "Integer"
        | TwoDecimal -> "TwoDecimal"
        | Percentage -> "Percentage"
        | Currency -> "Currency"
        | ShortDate -> "ShortDate"
        | DateAndTime -> "DateAndTime"
        | Custom code -> sprintf "Custom %s" (renderString code)

    let private renderCellProtection (p: CellProtection) : string =
        if p = CellProtection.Default then
            "CellProtection.Default"
        else
            [ if not p.Locked then yield "Locked = false"
              if p.Hidden then yield "Hidden = true" ]
            |> String.concat "; "
            |> sprintf "{ CellProtection.Default with %s }"

    let private renderCellStyle (s: CellStyle) : string =
        if s = CellStyle.Default then
            "CellStyle.Default"
        else
            [ if s.Font.IsSome then yield sprintf "Font = %s" (renderOption renderFontStyle s.Font)
              if s.Fill.IsSome then yield sprintf "Fill = %s" (renderOption renderFillStyle s.Fill)
              if s.Border.IsSome then yield sprintf "Border = %s" (renderOption renderBorderStyle s.Border)
              if s.NumberFormat.IsSome then
                  yield sprintf "NumberFormat = %s" (renderOption renderNumberFormat s.NumberFormat)
              if s.Alignment.IsSome then
                  yield sprintf "Alignment = %s" (renderOption renderAlignmentStyle s.Alignment)
              if s.Protection.IsSome then
                  yield sprintf "Protection = %s" (renderOption renderCellProtection s.Protection) ]
            |> String.concat "; "
            |> sprintf "{ CellStyle.Default with %s }"

    let private renderCellValue (v: CellValue) : string =
        match v with
        | Empty -> "Empty"
        | Text s -> sprintf "Text %s" (renderString s)
        | Number n -> sprintf "Number %s" (renderFloat n)
        | Boolean b -> sprintf "Boolean %s" (renderBool b)
        | Date d -> sprintf "Date %s" (renderDateTime d)
        | Formula(expr, cached) -> sprintf "Formula(%s, %s)" (renderString expr) (renderOption renderFloat cached)

    let private renderComparisonOperator (op: ComparisonOperator) : string =
        match op with
        | Equal -> "Equal"
        | NotEqual -> "NotEqual"
        | GreaterThan -> "GreaterThan"
        | LessThan -> "LessThan"
        | GreaterThanOrEqual -> "GreaterThanOrEqual"
        | LessThanOrEqual -> "LessThanOrEqual"
        | Between -> "Between"
        | NotBetween -> "NotBetween"

    let private renderConditionalFormatRule (rule: ConditionalFormatRule) : string =
        match rule with
        | CellValueRule(op, f1, f2, style) ->
            sprintf
                "CellValueRule(%s, %s, %s, %s)"
                (renderComparisonOperator op)
                (renderString f1)
                (renderOption renderString f2)
                (renderCellStyle style)
        | FormulaRule(f, style) -> sprintf "FormulaRule(%s, %s)" (renderString f) (renderCellStyle style)
        | ColorScale2(minColor, maxColor) -> sprintf "ColorScale2(%s, %s)" (renderColor minColor) (renderColor maxColor)
        | ColorScale3(minColor, midColor, maxColor) ->
            sprintf "ColorScale3(%s, %s, %s)" (renderColor minColor) (renderColor midColor) (renderColor maxColor)
        | DataBarRule color -> sprintf "DataBarRule(%s)" (renderColor color)
        | DuplicateValuesRule style -> sprintf "DuplicateValuesRule(%s)" (renderCellStyle style)
        | UniqueValuesRule style -> sprintf "UniqueValuesRule(%s)" (renderCellStyle style)

    let private renderValidationKind (kind: ValidationKind) : string =
        match kind with
        | ListValidation items -> items |> List.map renderString |> String.concat "; " |> sprintf "ListValidation [ %s ]"
        | ListFromRangeValidation(topLeft, bottomRight) ->
            sprintf "ListFromRangeValidation(%s, %s)" (renderCellRef topLeft) (renderCellRef bottomRight)
        | WholeNumberValidation(op, f1, f2) ->
            sprintf "WholeNumberValidation(%s, %s, %s)" (renderComparisonOperator op) (renderString f1) (renderOption renderString f2)
        | DecimalValidation(op, f1, f2) ->
            sprintf "DecimalValidation(%s, %s, %s)" (renderComparisonOperator op) (renderString f1) (renderOption renderString f2)
        | TextLengthValidation(op, f1, f2) ->
            sprintf "TextLengthValidation(%s, %s, %s)" (renderComparisonOperator op) (renderString f1) (renderOption renderString f2)
        | CustomValidation f -> sprintf "CustomValidation(%s)" (renderString f)

    let private renderErrorAlertStyle (s: ErrorAlertStyle) : string =
        match s with
        | Stop -> "Stop"
        | Warning -> "Warning"
        | Information -> "Information"

    let private renderHyperlinkTarget (t: HyperlinkTarget) : string =
        match t with
        | ExternalHyperlink url -> sprintf "ExternalHyperlink %s" (renderString url)
        | InternalHyperlink location -> sprintf "InternalHyperlink %s" (renderString location)

    /// `colOpt`/`value`/`style` mirror `CellEntry`'s own shape - `colOpt = None` means
    /// "the next column", same convention the DSL itself uses.
    let private renderCellEntry (colOpt: int option) (value: CellValue) (style: CellStyle option) : string =
        [ yield renderCellValue value
          match colOpt with
          | Some c -> yield sprintf "col = %d" c
          | None -> ()
          match style with
          | Some s -> yield sprintf "style = %s" (renderCellStyle s)
          | None -> () ]
        |> String.concat ", "
        |> sprintf "cell (%s)"

    let private renderRow (indexOpt: int option) (cellEntries: (int option * CellValue * CellStyle option) list) : string =
        let cellsStr =
            cellEntries |> List.map (fun (c, v, s) -> renderCellEntry c v s) |> String.concat "; "

        match indexOpt with
        | None -> sprintf "row [ %s ]" cellsStr
        | Some i -> sprintf "row ([ %s ], index = %d)" cellsStr i

    let private renderCommentEntry (c: CommentEntry) : string =
        if c.Author = "" then
            sprintf "comment (%s, %s)" (renderCellRef c.Cell) (renderString c.Text)
        else
            sprintf "comment (%s, %s, author = %s)" (renderCellRef c.Cell) (renderString c.Text) (renderString c.Author)

    let private renderHyperlinkEntry (h: HyperlinkEntry) : string =
        [ if h.TopLeft = h.BottomRight then
              yield renderCellRef h.TopLeft
          else
              yield renderCellRef h.TopLeft
              yield renderCellRef h.BottomRight
          yield renderHyperlinkTarget h.Target
          match h.Tooltip with
          | Some t -> yield sprintf "tooltip = %s" (renderString t)
          | None -> ()
          match h.Display with
          | Some d -> yield sprintf "display = %s" (renderString d)
          | None -> () ]
        |> String.concat ", "
        |> sprintf "hyperlink (%s)"

    let private renderDataValidationEntry (d: DataValidationEntry) : string =
        let a = d.Alert

        [ yield renderCellRef d.TopLeft
          yield renderCellRef d.BottomRight
          yield renderValidationKind d.Kind
          if not a.AllowBlank then yield "allowBlank = false"
          if a.ErrorStyle <> Stop then yield sprintf "errorStyle = %s" (renderErrorAlertStyle a.ErrorStyle)
          match a.ErrorTitle with
          | Some t -> yield sprintf "errorTitle = %s" (renderString t)
          | None -> ()
          match a.ErrorMessage with
          | Some t -> yield sprintf "errorMessage = %s" (renderString t)
          | None -> ()
          match a.InputTitle with
          | Some t -> yield sprintf "inputTitle = %s" (renderString t)
          | None -> ()
          match a.InputMessage with
          | Some t -> yield sprintf "inputMessage = %s" (renderString t)
          | None -> () ]
        |> String.concat ", "
        |> sprintf "dataValidation (%s)"

    /// No smart constructor exists for `SheetProtection` (see `Builders.fs`) - it's a plain
    /// record built the usual way, `{ SheetProtection.Default with ... }`.
    let private renderSheetProtection (p: SheetProtection) : string =
        let boolOpt name (v: bool option) =
            v |> Option.map (fun b -> sprintf "%s = %s" name (renderOption renderBool (Some b)))

        let parts =
            [ match p.Password with
              | Some pw -> yield sprintf "Password = %s" (renderOption renderString (Some pw))
              | None -> ()
              if p.Sheet <> true then yield sprintf "Sheet = %s" (renderBool p.Sheet)
              yield! boolOpt "Objects" p.Objects |> Option.toList
              yield! boolOpt "Scenarios" p.Scenarios |> Option.toList
              yield! boolOpt "FormatCells" p.FormatCells |> Option.toList
              yield! boolOpt "FormatColumns" p.FormatColumns |> Option.toList
              yield! boolOpt "FormatRows" p.FormatRows |> Option.toList
              yield! boolOpt "InsertColumns" p.InsertColumns |> Option.toList
              yield! boolOpt "InsertRows" p.InsertRows |> Option.toList
              yield! boolOpt "InsertHyperlinks" p.InsertHyperlinks |> Option.toList
              yield! boolOpt "DeleteColumns" p.DeleteColumns |> Option.toList
              yield! boolOpt "DeleteRows" p.DeleteRows |> Option.toList
              yield! boolOpt "SelectLockedCells" p.SelectLockedCells |> Option.toList
              yield! boolOpt "Sort" p.Sort |> Option.toList
              yield! boolOpt "AutoFilter" p.AutoFilter |> Option.toList
              yield! boolOpt "PivotTables" p.PivotTables |> Option.toList
              yield! boolOpt "SelectUnlockedCells" p.SelectUnlockedCells |> Option.toList ]

        if parts.IsEmpty then
            "SheetProtection.Default"
        else
            parts |> String.concat "; " |> sprintf "{ SheetProtection.Default with %s }"

    let private renderPageOrientation (o: PageOrientation) : string =
        match o with
        | Portrait -> "Portrait"
        | Landscape -> "Landscape"

    let private renderPaperSize (p: PaperSize) : string =
        match p with
        | Letter -> "Letter"
        | Legal -> "Legal"
        | Tabloid -> "Tabloid"
        | A3 -> "A3"
        | A4 -> "A4"
        | OtherPaperSize code -> sprintf "OtherPaperSize %d" code

    let private renderPrintScaling (s: PrintScaling) : string =
        match s with
        | ScalePercent pct -> sprintf "ScalePercent %d" pct
        | FitToPage(width, height) -> sprintf "FitToPage(%d, %d)" width height

    let private renderPageMargins (m: PageMargins) : string =
        if m = PageMargins.Default then
            "PageMargins.Default"
        else
            [ if m.Left <> PageMargins.Default.Left then yield sprintf "Left = %s" (renderFloat m.Left)
              if m.Right <> PageMargins.Default.Right then yield sprintf "Right = %s" (renderFloat m.Right)
              if m.Top <> PageMargins.Default.Top then yield sprintf "Top = %s" (renderFloat m.Top)
              if m.Bottom <> PageMargins.Default.Bottom then yield sprintf "Bottom = %s" (renderFloat m.Bottom)
              if m.Header <> PageMargins.Default.Header then yield sprintf "Header = %s" (renderFloat m.Header)
              if m.Footer <> PageMargins.Default.Footer then yield sprintf "Footer = %s" (renderFloat m.Footer) ]
            |> String.concat "; "
            |> sprintf "{ PageMargins.Default with %s }"

    let private renderPrintAreaRange ((topLeft, bottomRight): CellRef * CellRef) : string =
        sprintf "(%s, %s)" (renderCellRef topLeft) (renderCellRef bottomRight)

    let private renderPageSetup (ps: PageSetup) : string =
        if ps = PageSetup.Default then
            "PageSetup.Default"
        else
            [ if ps.Orientation <> PageSetup.Default.Orientation then
                  yield sprintf "Orientation = %s" (renderPageOrientation ps.Orientation)
              if ps.PaperSize.IsSome then yield sprintf "PaperSize = %s" (renderOption renderPaperSize ps.PaperSize)
              if ps.Scaling.IsSome then yield sprintf "Scaling = %s" (renderOption renderPrintScaling ps.Scaling)
              if ps.Margins <> PageMargins.Default then yield sprintf "Margins = %s" (renderPageMargins ps.Margins)
              if not ps.PrintArea.IsEmpty then
                  yield ps.PrintArea |> List.map renderPrintAreaRange |> String.concat "; " |> sprintf "PrintArea = [ %s ]"
              if ps.Header.IsSome then yield sprintf "Header = %s" (renderOption renderString ps.Header)
              if ps.Footer.IsSome then yield sprintf "Footer = %s" (renderOption renderString ps.Footer)
              if ps.EvenHeader.IsSome then yield sprintf "EvenHeader = %s" (renderOption renderString ps.EvenHeader)
              if ps.EvenFooter.IsSome then yield sprintf "EvenFooter = %s" (renderOption renderString ps.EvenFooter)
              if ps.FirstHeader.IsSome then yield sprintf "FirstHeader = %s" (renderOption renderString ps.FirstHeader)
              if ps.FirstFooter.IsSome then yield sprintf "FirstFooter = %s" (renderOption renderString ps.FirstFooter) ]
            |> String.concat "; "
            |> sprintf "{ PageSetup.Default with %s }"

    let private renderTableColumn (c: TableColumn) : string =
        sprintf "{ Name = %s; CalculatedFormula = %s }" (renderString c.Name) (renderOption renderString c.CalculatedFormula)

    let private renderTableStyle (s: TableStyle) : string =
        if s = TableStyle.Default then
            "TableStyle.Default"
        else
            [ if s.Name <> TableStyle.Default.Name then yield sprintf "Name = %s" (renderOption renderString s.Name)
              if s.ShowFirstColumn <> TableStyle.Default.ShowFirstColumn then
                  yield sprintf "ShowFirstColumn = %s" (renderBool s.ShowFirstColumn)
              if s.ShowLastColumn <> TableStyle.Default.ShowLastColumn then
                  yield sprintf "ShowLastColumn = %s" (renderBool s.ShowLastColumn)
              if s.ShowRowStripes <> TableStyle.Default.ShowRowStripes then
                  yield sprintf "ShowRowStripes = %s" (renderBool s.ShowRowStripes)
              if s.ShowColumnStripes <> TableStyle.Default.ShowColumnStripes then
                  yield sprintf "ShowColumnStripes = %s" (renderBool s.ShowColumnStripes) ]
            |> String.concat "; "
            |> sprintf "{ TableStyle.Default with %s }"

    /// No smart constructor exists for `TableEntry` (see `Builders.fs`) - it's a plain
    /// record built the usual way.
    let private renderTableEntry (t: TableEntry) : string =
        let columnsStr = t.Columns |> List.map renderTableColumn |> String.concat "; "

        sprintf
            "Table { TopLeft = %s; BottomRight = %s; Name = %s; Columns = [ %s ]; Style = %s }"
            (renderCellRef t.TopLeft)
            (renderCellRef t.BottomRight)
            (renderString t.Name)
            columnsStr
            (renderTableStyle t.Style)

    let private renderSparklineType (t: SparklineType) : string =
        match t with
        | Line -> "Line"
        | Column -> "Column"
        | WinLoss -> "WinLoss"

    let private renderSparklineStyle (s: SparklineStyle) : string =
        if s = SparklineStyle.Default then
            "SparklineStyle.Default"
        else
            [ if s.Type <> SparklineStyle.Default.Type then yield sprintf "Type = %s" (renderSparklineType s.Type)
              if s.Color.IsSome then yield sprintf "Color = %s" (renderOption renderColor s.Color)
              if s.LineWeight.IsSome then yield sprintf "LineWeight = %s" (renderOption renderFloat s.LineWeight)
              if s.ShowMarkers then yield "ShowMarkers = true"
              if s.ShowHigh then yield "ShowHigh = true"
              if s.ShowLow then yield "ShowLow = true"
              if s.ShowFirst then yield "ShowFirst = true"
              if s.ShowLast then yield "ShowLast = true"
              if s.ShowNegative then yield "ShowNegative = true" ]
            |> String.concat "; "
            |> sprintf "{ SparklineStyle.Default with %s }"

    let private renderSparklineCell (c: SparklineCell) : string =
        sprintf
            "{ Cell = %s; DataTopLeft = %s; DataBottomRight = %s }"
            (renderCellRef c.Cell)
            (renderCellRef c.DataTopLeft)
            (renderCellRef c.DataBottomRight)

    /// No smart constructor exists for `SparklineGroupEntry` (see `Builders.fs`) - it's a
    /// plain record built the usual way.
    let private renderSparklineGroupEntry (g: SparklineGroupEntry) : string =
        let sparklinesStr = g.Sparklines |> List.map renderSparklineCell |> String.concat "; "
        sprintf "SparklineGroup { Style = %s; Sparklines = [ %s ] }" (renderSparklineStyle g.Style) sparklinesStr

    let private renderDefinedNameScope (s: DefinedNameScope) : string =
        match s with
        | WorkbookScope -> "WorkbookScope"
        | SheetScope sheetName -> sprintf "SheetScope %s" (renderString sheetName)

    /// `definedName`/`sheetScopedDefinedName` always produce `Hidden = false` (see
    /// `Builders.fs`), so a hidden defined name - not exercised by any current example, but
    /// a real, round-trippable OOXML state - falls back to the raw record literal instead.
    let private renderDefinedNameEntry (d: DefinedNameEntry) : string =
        if not d.Hidden then
            match d.Scope with
            | WorkbookScope -> sprintf "definedName %s %s" (renderString d.Name) (renderString d.Formula)
            | SheetScope sheetName ->
                sprintf "sheetScopedDefinedName %s %s %s" (renderString sheetName) (renderString d.Name) (renderString d.Formula)
        else
            sprintf
                "{ Name = %s; Formula = %s; Scope = %s; Hidden = true }"
                (renderString d.Name)
                (renderString d.Formula)
                (renderDefinedNameScope d.Scope)

    /// No smart constructor exists for `WorkbookProtection` (see `Builders.fs`) - it's a
    /// plain record built the usual way, `{ WorkbookProtection.Default with ... }`.
    let private renderWorkbookProtection (p: WorkbookProtection) : string =
        let parts =
            [ match p.Password with
              | Some pw -> yield sprintf "Password = %s" (renderOption renderString (Some pw))
              | None -> ()
              match p.LockStructure with
              | Some v -> yield sprintf "LockStructure = %s" (renderOption renderBool (Some v))
              | None -> ()
              match p.LockWindows with
              | Some v -> yield sprintf "LockWindows = %s" (renderOption renderBool (Some v))
              | None -> () ]

        if parts.IsEmpty then
            "WorkbookProtection.Default"
        else
            parts |> String.concat "; " |> sprintf "{ WorkbookProtection.Default with %s }"

    /// Groups `ws.Cells` back into `row`/`cell` item source text, threading the same
    /// "next row"/"next column" cursor `SheetItems.cellsOf` folds over at interpretation
    /// time - so a row/cell only gets an explicit `index`/`col` where the source file
    /// actually has a gap, exactly mirroring how a human would write it by hand.
    let private renderWorksheetItems (ws: Worksheet) : string list =
        let rowItems =
            ws.Cells
            |> List.sortWith (fun a b -> CellRef.compare a.Ref b.Ref)
            |> List.groupBy (fun c -> c.Ref.Row)
            |> List.sortBy fst
            |> List.fold
                (fun (nextRow, acc) (rowIdx, cells) ->
                    let indexOpt = if rowIdx = nextRow then None else Some rowIdx

                    let cellEntries =
                        cells
                        |> List.sortBy (fun c -> c.Ref.Col)
                        |> List.fold
                            (fun (nextCol, acc2) c ->
                                let colOpt = if c.Ref.Col = nextCol then None else Some c.Ref.Col
                                (c.Ref.Col + 1, (colOpt, c.Value, c.Style) :: acc2))
                            (0, [])
                        |> snd
                        |> List.rev

                    (rowIdx + 1, renderRow indexOpt cellEntries :: acc))
                (0, [])
            |> snd
            |> List.rev

        let columnWidthItems =
            ws.ColumnProps
            |> Map.toList
            |> List.choose (fun (i, p) -> p.Width |> Option.map (fun w -> sprintf "ColumnWidth(%d, %s)" i (renderFloat w)))

        let rowHeightItems =
            ws.RowProps
            |> Map.toList
            |> List.choose (fun (i, p) -> p.Height |> Option.map (fun h -> sprintf "RowHeight(%d, %s)" i (renderFloat h)))

        let mergeItems =
            ws.MergedRanges
            |> List.map (fun m -> sprintf "Merge(%s, %s)" (renderCellRef m.TopLeft) (renderCellRef m.BottomRight))

        let freezeItems =
            ws.FreezePane |> Option.map (fun f -> sprintf "Freeze(%d, %d)" f.Rows f.Columns) |> Option.toList

        let autoFilterItems =
            ws.AutoFilter
            |> Option.map (fun a -> sprintf "autoFilter (%s, %s)" (renderCellRef a.TopLeft) (renderCellRef a.BottomRight))
            |> Option.toList

        let protectItems =
            ws.Protection |> Option.map (fun p -> sprintf "Protect(%s)" (renderSheetProtection p)) |> Option.toList

        let conditionalFormatItems =
            ws.ConditionalFormats
            |> List.map (fun e ->
                sprintf
                    "conditionalFormat (%s, %s, %s)"
                    (renderCellRef e.TopLeft)
                    (renderCellRef e.BottomRight)
                    (renderConditionalFormatRule e.Rule))

        let dataValidationItems = ws.DataValidations |> List.map renderDataValidationEntry
        let hyperlinkItems = ws.Hyperlinks |> List.map renderHyperlinkEntry
        let commentItems = ws.Comments |> List.map renderCommentEntry

        let pageSetupItems =
            ws.PageSetup |> Option.map (fun ps -> sprintf "PageSetup(%s)" (renderPageSetup ps)) |> Option.toList

        let tableItems = ws.Tables |> List.map renderTableEntry
        let sparklineGroupItems = ws.SparklineGroups |> List.map renderSparklineGroupEntry

        rowItems
        @ columnWidthItems
        @ rowHeightItems
        @ mergeItems
        @ freezeItems
        @ autoFilterItems
        @ protectItems
        @ conditionalFormatItems
        @ dataValidationItems
        @ hyperlinkItems
        @ commentItems
        @ pageSetupItems
        @ tableItems
        @ sparklineGroupItems

    /// Renders a whole `Workbook` as a self-contained F# script that rebuilds an
    /// equivalent file when run. `referenceLines` are whatever raw `#r` directives the
    /// caller needs to locate the SafeOpenXml assembly - this module has no opinion on
    /// that, since it depends entirely on where the script ends up living relative to the
    /// build output.
    let generate (referenceLines: string list) (outputFileName: string) (wb: Workbook) : string =
        let sb = StringBuilder()

        for line in referenceLines do
            sb.AppendLine(line: string) |> ignore

        sb.AppendLine() |> ignore
        sb.AppendLine("open SafeOpenXml") |> ignore
        sb.AppendLine("open type SafeOpenXml.SheetDsl") |> ignore
        sb.AppendLine() |> ignore

        let sheetVarNames = wb.Sheets |> List.mapi (fun i _ -> sprintf "sheet%d" i)

        List.zip sheetVarNames wb.Sheets
        |> List.iter (fun (varName, ws) ->
            let itemsStr = renderWorksheetItems ws |> String.concat "; "
            sb.AppendLine(sprintf "let %s = sheet %s [ %s ]" varName (renderString ws.Name) itemsStr) |> ignore)

        sb.AppendLine() |> ignore

        let sheetsListStr = sheetVarNames |> String.concat "; "

        let pipes =
            [ if not wb.DefinedNames.IsEmpty then
                  let namesStr = wb.DefinedNames |> List.map renderDefinedNameEntry |> String.concat "; "
                  yield sprintf "withDefinedNames [ %s ]" namesStr
              match wb.Protection with
              | Some p -> yield sprintf "withProtection (%s)" (renderWorkbookProtection p)
              | None -> () ]

        // Kept on one line, like every other generated statement in this function - see
        // `generate`'s own doc comment on why (sidesteps F#'s indentation-sensitive
        // offside rule entirely rather than trying to pretty-print a multi-line pipe).
        let wbExpr =
            (sprintf "workbook [ %s ]" sheetsListStr :: pipes) |> String.concat " |> "

        sb.AppendLine(sprintf "let wb = %s" wbExpr) |> ignore

        sb.AppendLine() |> ignore
        sb.AppendLine(sprintf "wb |> Workbook.save %s" (renderString outputFileName)) |> ignore

        sb.ToString()
