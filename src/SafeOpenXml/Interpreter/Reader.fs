namespace SafeOpenXml.Interpreter

open System
open System.Globalization
open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// The reverse transform - parses an existing OOXML `SpreadsheetDocument` back into the
/// DSL. Best-effort/semantic round-trip: values and formatting the DSL models are
/// preserved; OOXML features Core doesn't model (see MAPPING.md) are dropped rather than
/// causing a failure, so re-saving a foreign file loses only what Core never claimed to
/// represent.
module internal Reader =

    // See the matching comment in Writer.fs: these three would collide with this DSL's
    // own `SheetItem.SparklineGroup` case / `Spreadsheet.Formula`/`CellValue.Formula` if
    // their namespaces were opened directly, so they get type abbreviations instead.
    type X14SparklineGroup = DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup
    type X14Sparkline = DocumentFormat.OpenXml.Office2010.Excel.Sparkline
    type X14SparklineGroups = DocumentFormat.OpenXml.Office2010.Excel.SparklineGroups
    type X14SparklineTypeValues = DocumentFormat.OpenXml.Office2010.Excel.SparklineTypeValues

    let private inv = CultureInfo.InvariantCulture

    let inline private colorOf (c: ^T) : Color option =
        if isNull (box c) then
            None
        else
            let rgb = (^T: (member Rgb: HexBinaryValue) c)
            let indexed = (^T: (member Indexed: UInt32Value) c)
            let theme = (^T: (member Theme: UInt32Value) c)
            let tint = (^T: (member Tint: DoubleValue) c)

            if not (isNull (box rgb)) then
                let hex = rgb.Value
                let s = if hex.Length = 8 then hex.Substring(2) else hex
                let r = Convert.ToByte(s.Substring(0, 2), 16)
                let g = Convert.ToByte(s.Substring(2, 2), 16)
                let b = Convert.ToByte(s.Substring(4, 2), 16)
                Some(Rgb(r, g, b))
            elif not (isNull (box indexed)) then
                Some(Indexed(int indexed.Value))
            elif not (isNull (box theme)) then
                let t = if isNull (box tint) then None else Some tint.Value
                Some(Theme(int theme.Value, t))
            else
                None

    let inline private sideOf (side: ^T) : BorderSide option =
        if isNull (box side) then
            None
        else
            let styleVal = (^T: (member Style: EnumValue<BorderStyleValues>) side)

            if isNull (box styleVal) then
                None
            else
                let colorVal = (^T: (member Color: Spreadsheet.Color) side)
                Some { Style = BorderMapping.ofOpenXml styleVal.Value; Color = colorOf colorVal }

    let private fontOfOpenXml (f: Font) : FontStyle =
        { Name = f.FontName |> Option.ofObj |> Option.map (fun n -> n.Val.Value)
          Size = f.FontSize |> Option.ofObj |> Option.map (fun s -> s.Val.Value)
          Bold = not (isNull f.Bold)
          Italic = not (isNull f.Italic)
          Underline = not (isNull f.Underline)
          Strikethrough = not (isNull f.Strike)
          Color = colorOf f.Color }

    let private fillOfOpenXml (fill: Fill) : FillStyle option =
        match fill.PatternFill with
        | null -> None
        | pf when not (isNull pf.PatternType) && pf.PatternType.Value = PatternValues.Solid ->
            colorOf pf.ForegroundColor |> Option.map (fun c -> { Color = c })
        | _ -> None

    let private borderOfOpenXml (b: Border) : BorderStyle =
        { Left = sideOf b.LeftBorder
          Right = sideOf b.RightBorder
          Top = sideOf b.TopBorder
          Bottom = sideOf b.BottomBorder }

    let private alignmentOf (al: Alignment) : AlignmentStyle option =
        if isNull al then
            None
        else
            Some
                { Horizontal = al.Horizontal |> Option.ofObj |> Option.map (fun h -> AlignmentMapping.horizontalOfOpenXml h.Value)
                  Vertical = al.Vertical |> Option.ofObj |> Option.map (fun v -> AlignmentMapping.verticalOfOpenXml v.Value)
                  WrapText = not (isNull al.WrapText) && al.WrapText.Value }

    let private protectionOf (p: Protection) : CellProtection option =
        if isNull p then
            None
        else
            Some
                { Locked = isNull p.Locked || p.Locked.Value
                  Hidden = not (isNull p.Hidden) && p.Hidden.Value }

    // OOXML built-in numFmtIds Core doesn't have a named case for, but preserves as raw
    // format codes for round-trip fidelity instead of silently discarding them.
    let private otherBuiltinFormatCodes =
        Map.ofList
            [ 3u, "#,##0"
              4u, "#,##0.00"
              9u, "0%"
              37u, "#,##0 ;(#,##0)"
              38u, "#,##0 ;[Red](#,##0)"
              39u, "#,##0.00;(#,##0.00)"
              40u, "#,##0.00;[Red](#,##0.00)"
              49u, "@" ]

    let private numberFormatOf (numFmtId: uint32) (customFormats: Map<uint32, string>) : NumberFormat option =
        match numFmtId with
        | NumberFormatMapping.GeneralId -> None
        | NumberFormatMapping.IntegerId -> Some Integer
        | NumberFormatMapping.TwoDecimalId -> Some TwoDecimal
        | NumberFormatMapping.PercentageId -> Some Percentage
        | NumberFormatMapping.ShortDateId -> Some ShortDate
        | NumberFormatMapping.DateTimeId -> Some DateAndTime
        | id ->
            match customFormats |> Map.tryFind id with
            | Some code when code = NumberFormatMapping.CurrencyFormatCode -> Some Currency
            | Some code -> Some(Custom code)
            | None ->
                match otherBuiltinFormatCodes |> Map.tryFind id with
                | Some code -> Some(Custom code)
                | None -> None

    let private readCustomFormats (stylesheet: Stylesheet) : Map<uint32, string> =
        match stylesheet.NumberingFormats with
        | null -> Map.empty
        | nfs ->
            nfs.Elements<NumberingFormat>()
            |> Seq.map (fun nf -> nf.NumberFormatId.Value, nf.FormatCode.Value)
            |> Map.ofSeq

    let private cellStyleOf (stylesheet: Stylesheet) (customFormats: Map<uint32, string>) (styleIndex: uint32) : CellStyle option =
        match stylesheet.CellFormats with
        | null -> None
        | cellFormats ->
            match cellFormats.Elements<CellFormat>() |> Seq.tryItem (int styleIndex) with
            | None -> None
            | Some xf ->
                let font =
                    match xf.FontId with
                    | null -> None
                    | fid when fid.Value = 0u -> None
                    | fid -> stylesheet.Fonts.Elements<Font>() |> Seq.tryItem (int fid.Value) |> Option.map fontOfOpenXml

                let fill =
                    match xf.FillId with
                    | null -> None
                    | fillId when fillId.Value = 0u -> None
                    | fillId -> stylesheet.Fills.Elements<Fill>() |> Seq.tryItem (int fillId.Value) |> Option.bind fillOfOpenXml

                let border =
                    match xf.BorderId with
                    | null -> None
                    | borderId when borderId.Value = 0u -> None
                    | borderId -> stylesheet.Borders.Elements<Border>() |> Seq.tryItem (int borderId.Value) |> Option.map borderOfOpenXml

                let numFmt =
                    match xf.NumberFormatId with
                    | null -> None
                    | nid -> numberFormatOf nid.Value customFormats

                let alignment = alignmentOf xf.Alignment
                let protection = protectionOf xf.Protection

                let style =
                    { Font = font
                      Fill = fill
                      Border = border
                      NumberFormat = numFmt
                      Alignment = alignment
                      Protection = protection }

                if style = CellStyle.Default then None else Some style

    /// Parses a `dxf` (differential format) entry back into a `CellStyle`. Unlike
    /// `cellStyleOf`, everything is a direct child of the `dxf` element rather than an
    /// index into the shared Fonts/Fills/Borders/NumberingFormats collections, so there's
    /// no font/fill/border sub-registry to consult - and no null-guarding for granted, since
    /// (unlike `cellStyleOf`'s by-index lookups, which only ever see present elements)
    /// these direct child properties are `null` whenever that child is simply absent.
    let private dxfStyleOf (stylesheet: Stylesheet) (dxfId: uint32) : CellStyle option =
        match stylesheet.DifferentialFormats with
        | null -> None
        | dxfs ->
            match dxfs.Elements<DifferentialFormat>() |> Seq.tryItem (int dxfId) with
            | None -> None
            | Some dxf ->
                let font = if isNull dxf.Font then None else Some(fontOfOpenXml dxf.Font)
                let fill = if isNull dxf.Fill then None else fillOfOpenXml dxf.Fill
                let border = if isNull dxf.Border then None else Some(borderOfOpenXml dxf.Border)
                let alignment = alignmentOf dxf.Alignment

                let numFmt =
                    if isNull dxf.NumberingFormat then
                        None
                    else
                        let nf = dxf.NumberingFormat
                        let localFormats = Map.ofList [ nf.NumberFormatId.Value, nf.FormatCode.Value ]
                        numberFormatOf nf.NumberFormatId.Value localFormats

                let protection = protectionOf dxf.Protection

                let style =
                    { Font = font
                      Fill = fill
                      Border = border
                      NumberFormat = numFmt
                      Alignment = alignment
                      Protection = protection }

                if style = CellStyle.Default then None else Some style

    let private comparisonOperatorOf (op: ConditionalFormattingOperatorValues) : ComparisonOperator option =
        if op = ConditionalFormattingOperatorValues.Equal then Some Equal
        elif op = ConditionalFormattingOperatorValues.NotEqual then Some NotEqual
        elif op = ConditionalFormattingOperatorValues.GreaterThan then Some GreaterThan
        elif op = ConditionalFormattingOperatorValues.LessThan then Some LessThan
        elif op = ConditionalFormattingOperatorValues.GreaterThanOrEqual then Some GreaterThanOrEqual
        elif op = ConditionalFormattingOperatorValues.LessThanOrEqual then Some LessThanOrEqual
        elif op = ConditionalFormattingOperatorValues.Between then Some Between
        elif op = ConditionalFormattingOperatorValues.NotBetween then Some NotBetween
        // ContainsText/NotContains/BeginsWith/EndsWith are only used by rule kinds Core
        // doesn't parse (text/blank/error-contains rules) - never reached from CellIs.
        else None

    /// Parses a `cfRule` element back into a `ConditionalFormatRule`, or `None` for rule
    /// kinds Core doesn't model (icon sets, top/bottom N, text/blank/error-contains,
    /// above-average, time-period) - dropped rather than failing, per this module's
    /// best-effort philosophy.
    let private conditionalFormatRuleOf (stylesheet: Stylesheet option) (rule: ConditionalFormattingRule) : ConditionalFormatRule option =
        let styleOf () =
            match stylesheet, Option.ofObj rule.FormatId with
            | Some s, Some fid -> dxfStyleOf s fid.Value |> Option.defaultValue CellStyle.Default
            | _ -> CellStyle.Default

        let formulas = rule.Elements<Spreadsheet.Formula>() |> Seq.map (fun f -> f.Text) |> List.ofSeq

        if isNull rule.Type then
            None
        else
            let t = rule.Type.Value

            if t = ConditionalFormatValues.CellIs then
                match Option.ofObj rule.Operator |> Option.bind (fun o -> comparisonOperatorOf o.Value), formulas with
                | Some op, f1 :: rest -> Some(CellValueRule(op, f1, List.tryHead rest, styleOf ()))
                | _ -> None
            elif t = ConditionalFormatValues.Expression then
                match formulas with
                | f :: _ -> Some(FormulaRule(f, styleOf ()))
                | [] -> None
            elif t = ConditionalFormatValues.ColorScale then
                match rule.Elements<ColorScale>() |> Seq.tryHead with
                | None -> None
                | Some cs ->
                    let colors = cs.Elements<Spreadsheet.Color>() |> Seq.choose (fun c -> colorOf c) |> List.ofSeq

                    match colors with
                    | [ minC; maxC ] -> Some(ColorScale2(minC, maxC))
                    | [ minC; midC; maxC ] -> Some(ColorScale3(minC, midC, maxC))
                    | _ -> None
            elif t = ConditionalFormatValues.DataBar then
                rule.Elements<DataBar>()
                |> Seq.tryHead
                |> Option.bind (fun db -> db.Elements<Spreadsheet.Color>() |> Seq.tryHead)
                |> Option.bind (fun c -> colorOf c)
                |> Option.map DataBarRule
            elif t = ConditionalFormatValues.DuplicateValues then
                Some(DuplicateValuesRule(styleOf ()))
            elif t = ConditionalFormatValues.UniqueValues then
                Some(UniqueValuesRule(styleOf ()))
            else
                None

    let private dataValidationOperatorOf (op: DataValidationOperatorValues) : ComparisonOperator =
        if op = DataValidationOperatorValues.Equal then Equal
        elif op = DataValidationOperatorValues.NotEqual then NotEqual
        elif op = DataValidationOperatorValues.GreaterThan then GreaterThan
        elif op = DataValidationOperatorValues.LessThan then LessThan
        elif op = DataValidationOperatorValues.GreaterThanOrEqual then GreaterThanOrEqual
        elif op = DataValidationOperatorValues.LessThanOrEqual then LessThanOrEqual
        elif op = DataValidationOperatorValues.NotBetween then NotBetween
        else Between

    let private errorAlertStyleOf (s: DataValidationErrorStyleValues) : ErrorAlertStyle =
        if s = DataValidationErrorStyleValues.Warning then Warning
        elif s = DataValidationErrorStyleValues.Information then Information
        else Stop

    /// Reverses `Writer.listFormula`: strips the surrounding quotes and un-escapes doubled
    /// `"` characters. If `formula1` isn't a quoted literal, it's treated as a range
    /// reference (`ListFromRangeValidation`) instead - falling back to a single-item
    /// literal list only if it doesn't parse as a range either (e.g. a defined name,
    /// which isn't modeled).
    let private parseListFormula (formula1: string) : ValidationKind =
        if formula1.StartsWith("\"") && formula1.EndsWith("\"") && formula1.Length >= 2 then
            let inner = formula1.Substring(1, formula1.Length - 2)

            let items =
                inner.Replace("\"\"", "").Split(',')
                |> Array.map (fun s -> s.Replace("", "\""))
                |> List.ofArray

            ListValidation items
        else
            match formula1.Split(':') with
            | [| a; b |] ->
                try
                    ListFromRangeValidation(CellRef.ofA1 a, CellRef.ofA1 b)
                with _ ->
                    ListValidation [ formula1 ]
            | _ -> ListValidation [ formula1 ]

    let private dataValidationOf (dv: Spreadsheet.DataValidation) : DataValidationEntry option =
        match Option.ofObj dv.SequenceOfReferences |> Option.bind Seq.tryHead with
        | None -> None
        | Some refText ->
            let parts = refText.Value.Split(':')
            let topLeft = CellRef.ofA1 parts.[0]
            let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])
            let formula1 = if isNull dv.Formula1 then "" else dv.Formula1.Text
            let formula2 = if isNull dv.Formula2 then None else Some dv.Formula2.Text
            let operator () = Option.ofObj dv.Operator |> Option.map (fun o -> dataValidationOperatorOf o.Value) |> Option.defaultValue Between

            let kind =
                if isNull dv.Type then
                    CustomValidation formula1
                elif dv.Type.Value = DataValidationValues.List then
                    parseListFormula formula1
                elif dv.Type.Value = DataValidationValues.Whole then
                    WholeNumberValidation(operator (), formula1, formula2)
                elif dv.Type.Value = DataValidationValues.Decimal then
                    DecimalValidation(operator (), formula1, formula2)
                elif dv.Type.Value = DataValidationValues.TextLength then
                    TextLengthValidation(operator (), formula1, formula2)
                else
                    // Custom, and the unmodeled Date/Time/None types, all fall back to the
                    // raw formula text rather than being dropped.
                    CustomValidation formula1

            let alert =
                { AllowBlank = not (isNull dv.AllowBlank) && dv.AllowBlank.Value
                  ErrorStyle = Option.ofObj dv.ErrorStyle |> Option.map (fun s -> errorAlertStyleOf s.Value) |> Option.defaultValue Stop
                  ErrorTitle = dv.ErrorTitle |> Option.ofObj |> Option.map (fun s -> s.Value)
                  ErrorMessage = dv.Error |> Option.ofObj |> Option.map (fun s -> s.Value)
                  InputTitle = dv.PromptTitle |> Option.ofObj |> Option.map (fun s -> s.Value)
                  InputMessage = dv.Prompt |> Option.ofObj |> Option.map (fun s -> s.Value) }

            Some
                { TopLeft = topLeft
                  BottomRight = bottomRight
                  Kind = kind
                  Alert = alert }

    /// Parses a `hyperlink` element back into a `HyperlinkEntry`. An external target's
    /// `Id` is a relationship id, not a URL - the URL itself lives on the worksheet
    /// part's own `HyperlinkRelationships`, so `worksheetPart` is needed to look it up.
    let private hyperlinkOf (worksheetPart: WorksheetPart) (hl: Spreadsheet.Hyperlink) : HyperlinkEntry option =
        match Option.ofObj hl.Reference with
        | None -> None
        | Some refVal ->
            let parts = refVal.Value.Split(':')
            let topLeft = CellRef.ofA1 parts.[0]
            let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])
            let tooltip = hl.Tooltip |> Option.ofObj |> Option.map (fun t -> t.Value)
            let display = hl.Display |> Option.ofObj |> Option.map (fun d -> d.Value)

            let target =
                if not (isNull hl.Id) then
                    worksheetPart.HyperlinkRelationships
                    |> Seq.tryFind (fun r -> r.Id = hl.Id.Value)
                    |> Option.map (fun r -> ExternalHyperlink(r.Uri.OriginalString))
                elif not (isNull hl.Location) then
                    Some(InternalHyperlink(hl.Location.Value))
                else
                    None

            target
            |> Option.map (fun t ->
                { TopLeft = topLeft
                  BottomRight = bottomRight
                  Target = t
                  Tooltip = tooltip
                  Display = display })

    /// Parses one `comment` element back into a `CommentEntry`. The VML drawing part
    /// (the on-cell anchor/indicator) is never read - it isn't modeled data, just a
    /// rendering side-effect `Writer` reconstructs deterministically from `Cell` each
    /// time, so re-saving a round-tripped file regenerates a fresh one rather than
    /// preserving whatever exact VML the original had.
    let private commentOf (authors: string[]) (comment: Spreadsheet.Comment) : CommentEntry option =
        match Option.ofObj comment.Reference with
        | None -> None
        | Some refVal ->
            let author =
                match Option.ofObj comment.AuthorId with
                | Some aid when int aid.Value < authors.Length -> authors.[int aid.Value]
                | _ -> ""

            let text = if isNull comment.CommentText then "" else comment.CommentText.InnerText
            Some { Cell = CellRef.ofA1 refVal.Value; Author = author; Text = text }

    let private readComments (worksheetPart: WorksheetPart) : CommentEntry list =
        match worksheetPart.WorksheetCommentsPart with
        | null -> []
        | commentsPart ->
            let authors =
                match commentsPart.Comments.Authors with
                | null -> [||]
                | authorsEl -> authorsEl.Elements<Author>() |> Seq.map (fun a -> a.Text) |> Array.ofSeq

            match commentsPart.Comments.CommentList with
            | null -> []
            | commentList -> commentList.Elements<Spreadsheet.Comment>() |> Seq.choose (commentOf authors) |> List.ofSeq

    let private readSharedStrings (workbookPart: WorkbookPart) : string[] =
        match workbookPart.SharedStringTablePart with
        | null -> [||]
        | sstPart ->
            sstPart.SharedStringTable.Elements<SharedStringItem>()
            |> Seq.map (fun item -> item.InnerText)
            |> Array.ofSeq

    let private rawCellValueOf (sharedStrings: string[]) (c: Spreadsheet.Cell) : CellValue =
        let text () = if isNull c.CellValue then "" else c.CellValue.Text

        match c.CellFormula with
        | null ->
            let dataType = Option.ofObj c.DataType |> Option.map (fun d -> d.Value)

            if dataType = Some CellValues.SharedString then
                match Int32.TryParse(text (), NumberStyles.Integer, inv) with
                | true, idx when idx >= 0 && idx < sharedStrings.Length -> Text sharedStrings.[idx]
                | _ -> Text ""
            elif dataType = Some CellValues.Boolean then
                Boolean(text () = "1")
            elif dataType = Some CellValues.InlineString then
                match c.InlineString with
                | null -> Text ""
                | inlineStr -> Text inlineStr.InnerText
            elif dataType = Some CellValues.String then
                Text(text ())
            elif dataType = Some CellValues.Error then
                Text(text ())
            else
                if isNull c.CellValue then
                    Empty
                else
                    match Double.TryParse(text (), NumberStyles.Float, inv) with
                    | true, n -> Number n
                    | false, _ -> Empty
        | formula ->
            let cached =
                if isNull c.CellValue then
                    None
                else
                    match Double.TryParse(text (), NumberStyles.Float, inv) with
                    | true, n -> Some n
                    | false, _ -> None

            Formula(formula.Text, cached)

    let private readWorksheet
        (sharedStrings: string[])
        (stylesheet: Stylesheet option)
        (customFormats: Map<uint32, string>)
        (name: string)
        (printArea: (CellRef * CellRef) list)
        (worksheetPart: WorksheetPart)
        : Worksheet =
        let ws = worksheetPart.Worksheet
        let sheetData = ws.Elements<SheetData>() |> Seq.tryHead

        let cells =
            match sheetData with
            | None -> []
            | Some sd ->
                [ for row in sd.Elements<Row>() do
                      for c in row.Elements<Spreadsheet.Cell>() do
                          if not (isNull c.CellReference) && not (String.IsNullOrEmpty c.CellReference.Value) then
                              let cellRef = CellRef.ofA1 c.CellReference.Value
                              let styleIndex = if isNull c.StyleIndex then 0u else c.StyleIndex.Value
                              let style = stylesheet |> Option.bind (fun s -> cellStyleOf s customFormats styleIndex)
                              let rawValue = rawCellValueOf sharedStrings c

                              let value =
                                  match rawValue, style with
                                  | Number n, Some { NumberFormat = Some ShortDate }
                                  | Number n, Some { NumberFormat = Some DateAndTime } -> Date(DateTime.FromOADate n)
                                  | _ -> rawValue

                              yield { Ref = cellRef; Value = value; Style = style } ]

        let columnProps =
            ws.Elements<Columns>()
            |> Seq.collect (fun cols -> cols.Elements<Column>())
            |> Seq.collect (fun col ->
                let minC = int col.Min.Value - 1
                let maxC = int col.Max.Value - 1
                let width = if isNull col.Width then None else Some col.Width.Value
                [ for c in minC..maxC -> c, { Width = width } ])
            |> Map.ofSeq

        let rowProps =
            match sheetData with
            | None -> Map.empty
            | Some sd ->
                sd.Elements<Row>()
                |> Seq.choose (fun row ->
                    if isNull row.Height then
                        None
                    else
                        let idx = int row.RowIndex.Value - 1
                        Some(idx, { Height = Some row.Height.Value }))
                |> Map.ofSeq

        let mergedRanges : MergedRange list =
            ws.Elements<MergeCells>()
            |> Seq.collect (fun mc -> mc.Elements<MergeCell>())
            |> Seq.map (fun mc ->
                let parts = mc.Reference.Value.Split(':')

                { TopLeft = CellRef.ofA1 parts.[0]
                  BottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0]) }
                : MergedRange)
            |> List.ofSeq

        let autoFilter : AutoFilterRange option =
            ws.Elements<Spreadsheet.AutoFilter>()
            |> Seq.tryHead
            |> Option.bind (fun af -> Option.ofObj af.Reference)
            |> Option.map (fun refVal ->
                let parts = refVal.Value.Split(':')

                { TopLeft = CellRef.ofA1 parts.[0]
                  BottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0]) })

        // Password never round-trips: the hash isn't reversible, so a protected file read
        // back through the DSL loses password enforcement unless it's re-supplied - see
        // SheetProtection's own doc comment.
        let sheetProtection =
            ws.Elements<Spreadsheet.SheetProtection>()
            |> Seq.tryHead
            |> Option.map (fun sp ->
                let flag (v: BooleanValue) = if isNull v then None else Some v.Value

                { Password = None
                  Sheet = not (isNull sp.Sheet) && sp.Sheet.Value
                  Objects = flag sp.Objects
                  Scenarios = flag sp.Scenarios
                  FormatCells = flag sp.FormatCells
                  FormatColumns = flag sp.FormatColumns
                  FormatRows = flag sp.FormatRows
                  InsertColumns = flag sp.InsertColumns
                  InsertRows = flag sp.InsertRows
                  InsertHyperlinks = flag sp.InsertHyperlinks
                  DeleteColumns = flag sp.DeleteColumns
                  DeleteRows = flag sp.DeleteRows
                  SelectLockedCells = flag sp.SelectLockedCells
                  Sort = flag sp.Sort
                  AutoFilter = flag sp.AutoFilter
                  PivotTables = flag sp.PivotTables
                  SelectUnlockedCells = flag sp.SelectUnlockedCells })

        let freezePane =
            ws.Elements<SheetViews>()
            |> Seq.collect (fun svs -> svs.Elements<SheetView>())
            |> Seq.collect (fun sv -> sv.Elements<Pane>())
            |> Seq.tryHead
            |> Option.bind (fun pane ->
                if not (isNull pane.State) && pane.State.Value = PaneStateValues.Frozen then
                    let rows = if isNull pane.VerticalSplit then 0 else int pane.VerticalSplit.Value
                    let cols = if isNull pane.HorizontalSplit then 0 else int pane.HorizontalSplit.Value
                    Some { Rows = rows; Columns = cols }
                else
                    None)

        let conditionalFormats =
            ws.Elements<ConditionalFormatting>()
            |> Seq.collect (fun cf ->
                match Option.ofObj cf.SequenceOfReferences |> Option.bind Seq.tryHead with
                | None -> Seq.empty
                | Some refText ->
                    let parts = refText.Value.Split(':')
                    let topLeft = CellRef.ofA1 parts.[0]
                    let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])

                    cf.Elements<ConditionalFormattingRule>()
                    |> Seq.choose (fun r -> conditionalFormatRuleOf stylesheet r)
                    |> Seq.map (fun rule -> { TopLeft = topLeft; BottomRight = bottomRight; Rule = rule }))
            |> List.ofSeq

        let dataValidations =
            ws.Elements<DataValidations>()
            |> Seq.collect (fun dvs -> dvs.Elements<Spreadsheet.DataValidation>())
            |> Seq.choose dataValidationOf
            |> List.ofSeq

        let hyperlinks =
            ws.Elements<Hyperlinks>()
            |> Seq.collect (fun hls -> hls.Elements<Spreadsheet.Hyperlink>())
            |> Seq.choose (fun hl -> hyperlinkOf worksheetPart hl)
            |> List.ofSeq

        // `fitToPage` lives on `sheetPr/pageSetUpPr`, a sibling of `pageSetup` rather than
        // an attribute on it - needed to tell `ScalePercent`/`FitToPage` apart, since both
        // `scale` and `fitToWidth`/`fitToHeight` can be present in the file regardless of
        // which one Excel is actually honoring (see `Writer.pageSetupElement`).
        let fitToPageFlag =
            ws.Elements<SheetProperties>()
            |> Seq.tryHead
            |> Option.bind (fun sp -> Option.ofObj sp.PageSetupProperties)
            |> Option.bind (fun psp -> Option.ofObj psp.FitToPage)
            |> Option.map (fun v -> v.Value)
            |> Option.defaultValue false

        let pageSetup =
            ws.Elements<Spreadsheet.PageSetup>()
            |> Seq.tryHead
            |> Option.map (fun setup ->
                let orientation =
                    setup.Orientation |> Option.ofObj |> Option.map (fun o -> OrientationMapping.ofOpenXml o.Value) |> Option.defaultValue Portrait

                let paperSize = setup.PaperSize |> Option.ofObj |> Option.map (fun p -> PaperSizeMapping.ofOpenXml p.Value)

                let scaling =
                    if fitToPageFlag then
                        let width = setup.FitToWidth |> Option.ofObj |> Option.map (fun v -> int v.Value) |> Option.defaultValue 1
                        let height = setup.FitToHeight |> Option.ofObj |> Option.map (fun v -> int v.Value) |> Option.defaultValue 1
                        Some(FitToPage(width, height))
                    else
                        setup.Scale |> Option.ofObj |> Option.map (fun v -> ScalePercent(int v.Value))

                let margins =
                    ws.Elements<Spreadsheet.PageMargins>()
                    |> Seq.tryHead
                    |> Option.map (fun m ->
                        { Left = m.Left.Value
                          Right = m.Right.Value
                          Top = m.Top.Value
                          Bottom = m.Bottom.Value
                          Header = m.Header.Value
                          Footer = m.Footer.Value })
                    |> Option.defaultValue PageMargins.Default

                // `OddHeader`/`EvenHeader`/`FirstHeader`/etc. all derive from the same
                // `XstringType` base (a plain text-content element), so one helper covers
                // all six rather than repeating the same null-check/`.Text` projection.
                let textElementOf (e: XstringType) : string option =
                    if isNull (box e) then None else Some e.Text

                let headerFooter = ws.Elements<HeaderFooter>() |> Seq.tryHead
                let textOf (project: HeaderFooter -> XstringType) = headerFooter |> Option.bind (fun hf -> textElementOf (project hf))

                { Orientation = orientation
                  PaperSize = paperSize
                  Scaling = scaling
                  Margins = margins
                  PrintArea = []
                  Header = textOf (fun hf -> hf.OddHeader)
                  Footer = textOf (fun hf -> hf.OddFooter)
                  EvenHeader = textOf (fun hf -> hf.EvenHeader)
                  EvenFooter = textOf (fun hf -> hf.EvenFooter)
                  FirstHeader = textOf (fun hf -> hf.FirstHeader)
                  FirstFooter = textOf (fun hf -> hf.FirstFooter) })

        // `PrintArea` is threaded in from `load`, which resolves the workbook-level
        // `_xlnm.Print_Area` defined name before any worksheet is parsed - see that
        // function and `PageSetup`'s own doc comment for why.
        let pageSetup =
            match pageSetup, printArea with
            | Some ps, [] -> Some ps
            | Some ps, area -> Some { ps with PrintArea = area }
            | None, [] -> None
            | None, area -> Some { PageSetup.Default with PrintArea = area }

        // Core never writes a totals row (see MAPPING.md), and drops one on read too -
        // `TableColumn.TotalsRowFunction`/`.TotalsRowLabel` are simply never consulted
        // here, same "unmodeled state is dropped, not failed on" philosophy as everywhere
        // else in this module.
        let tableColumnOf (tc: Spreadsheet.TableColumn) : TableColumn =
            { Name = if isNull tc.Name then "" else tc.Name.Value
              CalculatedFormula = tc.CalculatedColumnFormula |> Option.ofObj |> Option.map (fun f -> f.Text) }

        let tableStyleOf (tsi: TableStyleInfo) : TableStyle =
            if isNull tsi then
                TableStyle.Default
            else
                let flag (v: BooleanValue) = not (isNull v) && v.Value

                { Name = tsi.Name |> Option.ofObj |> Option.map (fun n -> n.Value)
                  ShowFirstColumn = flag tsi.ShowFirstColumn
                  ShowLastColumn = flag tsi.ShowLastColumn
                  ShowRowStripes = flag tsi.ShowRowStripes
                  ShowColumnStripes = flag tsi.ShowColumnStripes }

        let tableOf (table: Spreadsheet.Table) : TableEntry option =
            match Option.ofObj table.Reference, Option.ofObj table.Name with
            | Some refVal, Some nameVal ->
                let parts = refVal.Value.Split(':')
                let topLeft = CellRef.ofA1 parts.[0]
                let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])

                let columns =
                    match table.TableColumns with
                    | null -> []
                    | cols -> cols.Elements<Spreadsheet.TableColumn>() |> Seq.map tableColumnOf |> List.ofSeq

                Some
                    { TopLeft = topLeft
                      BottomRight = bottomRight
                      Name = nameVal.Value
                      Columns = columns
                      Style = tableStyleOf table.TableStyleInfo }
            | _ -> None

        let tables =
            ws.Elements<TableParts>()
            |> Seq.collect (fun tps -> tps.Elements<TablePart>())
            |> Seq.choose (fun tp ->
                match Option.ofObj tp.Id with
                | None -> None
                | Some relId ->
                    match worksheetPart.GetPartById(relId.Value) with
                    | :? TableDefinitionPart as tdp -> tableOf tdp.Table
                    | _ -> None)
            |> List.ofSeq

        let charts, images = DrawingReader.readChartsAndImages worksheetPart ws

        // Same sheet-qualifier stripping as `parsePrintAreaFormula` - a sparkline's data
        // range is always on the same sheet as the sparkline itself in what Core writes,
        // and this doesn't model a foreign file's sparklines pointing elsewhere.
        let stripSheetQualifier (s: string) =
            match s.LastIndexOf('!') with
            | -1 -> s
            | i -> s.Substring(i + 1)

        let sparklineCellOf (sl: X14Sparkline) : SparklineCell option =
            match Option.ofObj sl.Formula, Option.ofObj sl.ReferenceSequence with
            | Some f, Some r ->
                try
                    let dataParts = (stripSheetQualifier f.Text).Split(':')
                    let dataTopLeft = CellRef.ofA1 dataParts.[0]
                    let dataBottomRight = CellRef.ofA1 (if dataParts.Length > 1 then dataParts.[1] else dataParts.[0])
                    let cell = CellRef.ofA1 ((stripSheetQualifier r.Text).Split(':').[0])
                    Some { Cell = cell; DataTopLeft = dataTopLeft; DataBottomRight = dataBottomRight }
                with _ ->
                    None
            | _ -> None

        // Core never models the richer axis/date-axis/per-role-color settings - see
        // MAPPING.md - so those are simply never consulted here, same philosophy as the
        // rest of this module.
        let sparklineGroupEntryOf (group: X14SparklineGroup) : SparklineGroupEntry =
            let flag (v: BooleanValue) = not (isNull v) && v.Value

            let style =
                { Type =
                    Option.ofObj group.Type
                    |> Option.map (fun t -> SparklineMapping.ofOpenXml t.Value)
                    |> Option.defaultValue Line
                  Color = colorOf group.SeriesColor
                  LineWeight = group.LineWeight |> Option.ofObj |> Option.map (fun v -> v.Value)
                  ShowMarkers = flag group.Markers
                  ShowHigh = flag group.High
                  ShowLow = flag group.Low
                  ShowFirst = flag group.First
                  ShowLast = flag group.Last
                  ShowNegative = flag group.Negative }

            let sparklines =
                match group.Sparklines with
                | null -> []
                | sls -> sls.Elements<X14Sparkline>() |> Seq.choose sparklineCellOf |> List.ofSeq

            { Style = style; Sparklines = sparklines }

        let sparklineGroups =
            ws.Elements<WorksheetExtensionList>()
            |> Seq.collect (fun extList -> extList.Elements<WorksheetExtension>())
            |> Seq.collect (fun ext -> ext.Elements<X14SparklineGroups>())
            |> Seq.collect (fun sgs -> sgs.Elements<X14SparklineGroup>())
            |> Seq.map sparklineGroupEntryOf
            |> List.ofSeq

        let comments = readComments worksheetPart

        { Name = name
          Cells = cells
          ColumnProps = columnProps
          RowProps = rowProps
          MergedRanges = mergedRanges
          FreezePane = freezePane
          AutoFilter = autoFilter
          Protection = sheetProtection
          ConditionalFormats = conditionalFormats
          DataValidations = dataValidations
          Hyperlinks = hyperlinks
          Comments = comments
          PageSetup = pageSetup
          Tables = tables
          SparklineGroups = sparklineGroups
          Charts = charts
          Images = images }

    /// Reverses `Writer.absolutePrintAreaRange`/`printAreaDefinedNameElement`: splits the
    /// comma-separated range list, strips each range's sheet-name qualifier (print area is
    /// already inherently per-sheet, so it's discarded rather than checked against the
    /// current sheet's own name) and `$` markers, then parses the bare `"A1:D10"` range.
    /// Accepts both quoted (`'Sheet 1'!...`) and unquoted (`Sheet1!...`) qualifiers, since
    /// a foreign file may use either. A range this can't parse is dropped, not failed on -
    /// same philosophy as the rest of this module.
    let private parsePrintAreaFormula (formula: string) : (CellRef * CellRef) list =
        formula.Split(',')
        |> Array.choose (fun part ->
            let refPart =
                match part.LastIndexOf('!') with
                | -1 -> part
                | i -> part.Substring(i + 1)

            let cellParts = refPart.Replace("$", "").Split(':')

            try
                let topLeft = CellRef.ofA1 cellParts.[0]
                let bottomRight = CellRef.ofA1 (if cellParts.Length > 1 then cellParts.[1] else cellParts.[0])
                Some(topLeft, bottomRight)
            with _ ->
                None)
        |> List.ofArray

    let load (document: SpreadsheetDocument) : Workbook =
        let workbookPart = document.WorkbookPart
        let sharedStrings = readSharedStrings workbookPart

        let stylesheet =
            match workbookPart.WorkbookStylesPart with
            | null -> None
            | sp -> Some sp.Stylesheet

        let customFormats =
            stylesheet |> Option.map readCustomFormats |> Option.defaultValue Map.empty

        let sheetEls = workbookPart.Workbook.Sheets.Elements<Sheet>() |> List.ofSeq
        let sheetNames = sheetEls |> List.map (fun s -> s.Name.Value) |> Array.ofList

        let allDefinedNameEls =
            match workbookPart.Workbook.DefinedNames with
            | null -> []
            | dns -> dns.Elements<Spreadsheet.DefinedName>() |> List.ofSeq

        // Print areas are represented as a reserved, hidden, sheet-scoped defined name
        // (`_xlnm.Print_Area`) rather than an ordinary `DefinedNameEntry` - resolved into
        // each sheet's `PageSetup.PrintArea` here, before any worksheet is parsed, so
        // `readWorksheet` can just fold it into the `PageSetup` it builds from the
        // worksheet's own elements. See `PageSetup`'s own doc comment for why.
        let printAreaBySheetIndex : Map<int, (CellRef * CellRef) list> =
            allDefinedNameEls
            |> List.choose (fun dn ->
                match Option.ofObj dn.Name, Option.ofObj dn.LocalSheetId with
                | Some nameVal, Some idx when nameVal.Value = "_xlnm.Print_Area" -> Some(int idx.Value, parsePrintAreaFormula dn.Text)
                | _ -> None)
            |> Map.ofList

        let definedNames =
            allDefinedNameEls
            |> List.choose (fun dn ->
                match Option.ofObj dn.Name with
                | None -> None
                | Some nameVal when nameVal.Value = "_xlnm.Print_Area" -> None
                | Some nameVal ->
                    let scope =
                        match Option.ofObj dn.LocalSheetId with
                        | Some idx when int idx.Value < sheetNames.Length -> SheetScope sheetNames.[int idx.Value]
                        | _ -> WorkbookScope

                    Some
                        { Name = nameVal.Value
                          Formula = dn.Text
                          Scope = scope
                          Hidden = not (isNull dn.Hidden) && dn.Hidden.Value })

        let sheets =
            sheetEls
            |> List.mapi (fun i sheetEl ->
                let worksheetPart = workbookPart.GetPartById(sheetEl.Id.Value) :?> WorksheetPart
                let printArea = printAreaBySheetIndex |> Map.tryFind i |> Option.defaultValue []
                readWorksheet sharedStrings stylesheet customFormats sheetEl.Name.Value printArea worksheetPart)

        // Password never round-trips here either, same reason as `SheetProtection`'s own
        // doc comment: the hash isn't reversible.
        let protection =
            workbookPart.Workbook.WorkbookProtection
            |> Option.ofObj
            |> Option.map (fun wp ->
                let flag (v: BooleanValue) = if isNull v then None else Some v.Value

                { Password = None
                  LockStructure = flag wp.LockStructure
                  LockWindows = flag wp.LockWindows })

        { Sheets = sheets
          DefinedNames = definedNames
          Protection = protection }

    let loadFromStream (stream: Stream) : Workbook =
        use document = SpreadsheetDocument.Open(stream, false)
        load document

    let loadFromFile (path: string) : Workbook =
        use document = SpreadsheetDocument.Open(path, false)
        load document
