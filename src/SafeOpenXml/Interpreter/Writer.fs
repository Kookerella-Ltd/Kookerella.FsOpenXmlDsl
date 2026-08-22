namespace SafeOpenXml.Interpreter

open System
open System.Collections.Generic
open System.Globalization
open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// Interprets the DSL - compiles a `Workbook` value into OOXML calls against the
/// DocumentFormat.OpenXml SDK.
module internal Writer =

    let private inv = CultureInfo.InvariantCulture

    let private rangeReference (topLeft: CellRef) (bottomRight: CellRef) : string =
        sprintf "%s:%s" (CellRef.toA1 topLeft) (CellRef.toA1 bottomRight)

    /// The classic Excel worksheet-password hash (widely documented, e.g. [MS-OFFCRYPTO]
    /// §2.3.3, and reproduced identically across independent third-party spreadsheet
    /// libraries) - a weak XOR/rotate checksum, not real security, chosen over the newer
    /// salted SHA-512 scheme for the broadest possible Excel-version compatibility.
    /// Passwords longer than 15 characters are truncated, matching Excel's own limit for
    /// this scheme.
    let private legacyPasswordHash (password: string) : string =
        let truncated = if password.Length > 15 then password.Substring(0, 15) else password
        let mutable hash = 0

        for i in truncated.Length - 1 .. -1 .. 0 do
            hash <- ((hash <<< 1) &&& 0x7FFF) ||| ((hash >>> 14) &&& 0x01)
            hash <- hash ^^^ int truncated.[i]

        hash <- ((hash <<< 1) &&& 0x7FFF) ||| ((hash >>> 14) &&& 0x01)
        hash <- hash ^^^ truncated.Length
        hash <- hash ^^^ 0xCE4B
        sprintf "%04X" hash

    /// Builds a `sheetProtection` element - a thin, direct pass-through of `SheetProtection`
    /// (see that type's own doc comment for why: several of these flags aren't "true means
    /// allowed", and guessing a default backwards would be a silent, schema-valid, but
    /// wrong-behavior bug). Only `Sheet` is always written; every other flag is written
    /// only when the caller gave it explicitly, leaving Excel's own default to apply
    /// otherwise.
    let private sheetProtectionElement (sp: SheetProtection) : Spreadsheet.SheetProtection =
        let el = Spreadsheet.SheetProtection(Sheet = BooleanValue(sp.Sheet))
        sp.Password |> Option.iter (fun pwd -> el.Password <- HexBinaryValue(legacyPasswordHash pwd))
        sp.Objects |> Option.iter (fun v -> el.Objects <- BooleanValue(v))
        sp.Scenarios |> Option.iter (fun v -> el.Scenarios <- BooleanValue(v))
        sp.FormatCells |> Option.iter (fun v -> el.FormatCells <- BooleanValue(v))
        sp.FormatColumns |> Option.iter (fun v -> el.FormatColumns <- BooleanValue(v))
        sp.FormatRows |> Option.iter (fun v -> el.FormatRows <- BooleanValue(v))
        sp.InsertColumns |> Option.iter (fun v -> el.InsertColumns <- BooleanValue(v))
        sp.InsertRows |> Option.iter (fun v -> el.InsertRows <- BooleanValue(v))
        sp.InsertHyperlinks |> Option.iter (fun v -> el.InsertHyperlinks <- BooleanValue(v))
        sp.DeleteColumns |> Option.iter (fun v -> el.DeleteColumns <- BooleanValue(v))
        sp.DeleteRows |> Option.iter (fun v -> el.DeleteRows <- BooleanValue(v))
        sp.SelectLockedCells |> Option.iter (fun v -> el.SelectLockedCells <- BooleanValue(v))
        sp.Sort |> Option.iter (fun v -> el.Sort <- BooleanValue(v))
        sp.AutoFilter |> Option.iter (fun v -> el.AutoFilter <- BooleanValue(v))
        sp.PivotTables |> Option.iter (fun v -> el.PivotTables <- BooleanValue(v))
        sp.SelectUnlockedCells |> Option.iter (fun v -> el.SelectUnlockedCells <- BooleanValue(v))
        el

    /// Builds a `definedName` element. `sheetIndex` maps sheet name to its 0-based
    /// position for `SheetScope`'s translation to OOXML's `localSheetId`; a `SheetScope`
    /// naming a sheet not in this workbook is a genuine caller mistake, not something to
    /// paper over, so it raises rather than silently dropping the name.
    let private definedNameElement (sheetIndex: Map<string, int>) (entry: DefinedNameEntry) : Spreadsheet.DefinedName =
        let dn = Spreadsheet.DefinedName(entry.Formula)
        dn.Name <- StringValue(entry.Name)
        if entry.Hidden then dn.Hidden <- BooleanValue(true)

        match entry.Scope with
        | WorkbookScope -> ()
        | SheetScope sheetName ->
            match sheetIndex |> Map.tryFind sheetName with
            | Some idx -> dn.LocalSheetId <- UInt32Value(uint32 idx)
            | None ->
                invalidArg
                    (nameof entry)
                    (sprintf "DefinedName '%s' is scoped to sheet '%s', which isn't in this workbook" entry.Name sheetName)

        dn

    /// Same as `rangeReference`, but a single-cell range (`TopLeft = BottomRight`) writes
    /// as just `"A1"` rather than `"A1:A1"`, matching how Excel itself writes a
    /// single-cell hyperlink's `ref` attribute.
    let private hyperlinkRangeReference (topLeft: CellRef) (bottomRight: CellRef) : string =
        if topLeft = bottomRight then CellRef.toA1 topLeft else rangeReference topLeft bottomRight

    /// Builds a `hyperlink` element. External targets need a relationship on the
    /// worksheet part (the `r:id` that `.Id` refers to); internal (same-workbook)
    /// targets just go straight into `.Location`, no relationship involved.
    let private hyperlinkElement (worksheetPart: WorksheetPart) (entry: HyperlinkEntry) : Spreadsheet.Hyperlink =
        let hl = Spreadsheet.Hyperlink(Reference = StringValue(hyperlinkRangeReference entry.TopLeft entry.BottomRight))
        entry.Tooltip |> Option.iter (fun t -> hl.Tooltip <- StringValue(t))
        entry.Display |> Option.iter (fun d -> hl.Display <- StringValue(d))

        match entry.Target with
        | ExternalHyperlink url ->
            let relationship = worksheetPart.AddHyperlinkRelationship(Uri(url, UriKind.RelativeOrAbsolute), true)
            hl.Id <- StringValue(relationship.Id)
        | InternalHyperlink location -> hl.Location <- StringValue(location)

        hl

    /// Builds the `comments1.xml`-equivalent `Comments` root: a deduplicated `Authors`
    /// list (interned in order of first appearance, same dedup shape as shared strings)
    /// plus one `comment` per entry referencing its author by index.
    let private commentsRoot (entries: CommentEntry list) : Comments =
        let authors = ResizeArray<string>()
        let authorIndex = Dictionary<string, int>()

        let internAuthor (name: string) =
            match authorIndex.TryGetValue name with
            | true, idx -> idx
            | false, _ ->
                let idx = authors.Count
                authors.Add name
                authorIndex.[name] <- idx
                idx

        let commentList = CommentList()

        entries
        |> List.iter (fun entry ->
            let authorId = internAuthor entry.Author
            let commentText = CommentText(Text = Spreadsheet.Text(entry.Text))

            // No `shapeId` attribute here - the SpreadsheetML schema doesn't declare one
            // on `comment` (unlike some other OOXML comment-ish elements). The link
            // between a comment and its VML shape is purely by row/column match in the
            // VML's own `x:ClientData` (`vmlDrawingContent`), not an id on this element.
            let comment =
                Spreadsheet.Comment(
                    Reference = StringValue(CellRef.toA1 entry.Cell),
                    AuthorId = UInt32Value(uint32 authorId),
                    CommentText = commentText
                )

            commentList.AppendChild(comment) |> ignore)

        let authorsEl = Authors()
        authors |> Seq.iter (fun name -> authorsEl.AppendChild(Author(name)) |> ignore)

        Comments(Authors = authorsEl, CommentList = commentList)

    /// Builds the accompanying legacy VML drawing content Excel pairs with `Comments` for
    /// the on-cell red-triangle indicator and the comment box's hover position. Only the
    /// per-comment position/row/column varies - the wrapping namespaces, `shapelayout`,
    /// and `shapetype` are the same fixed boilerplate Excel itself always emits, so
    /// they're a template rather than built from typed elements (this is the one place in
    /// the interpreter that isn't - VML predates OOXML's schema-driven object model, and
    /// nothing here carries user-controlled text that would need escaping).
    let private vmlDrawingContent (entries: CommentEntry list) : string =
        let shape (idx: int) (entry: CommentEntry) : string =
            let shapeId = 1025 + idx
            let row = entry.Cell.Row
            let col = entry.Cell.Col
            let leftPt = float (col + 1) * 60.0
            let topPt = float row * 15.0

            sprintf
                "<v:shape id=\"_x0000_s%d\" type=\"#_x0000_t202\" style='position:absolute;margin-left:%gpt;margin-top:%gpt;width:108pt;height:59.25pt;z-index:%d;visibility:hidden' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n\
                 <v:fill color2=\"#ffffe1\"/>\n\
                 <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n\
                 <v:path o:connecttype=\"none\"/>\n\
                 <v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>\n\
                 <x:ClientData ObjectType=\"Note\">\n\
                 <x:MoveWithCells/>\n\
                 <x:SizeWithCells/>\n\
                 <x:Anchor>%d, 15, %d, 2, %d, 31, %d, 4</x:Anchor>\n\
                 <x:AutoFill>False</x:AutoFill>\n\
                 <x:Row>%d</x:Row>\n\
                 <x:Column>%d</x:Column>\n\
                 </x:ClientData>\n\
                 </v:shape>"
                shapeId
                leftPt
                topPt
                (idx + 1)
                (col + 1)
                row
                (col + 3)
                (row + 4)
                row
                col

        let shapes = entries |> List.mapi shape |> String.concat "\n"

        sprintf
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n\
             <o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>\n\
             <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">\n\
             <v:stroke joinstyle=\"miter\"/>\n\
             <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n\
             </v:shapetype>\n\
             %s\n\
             </xml>"
            shapes

    let private comparisonOperatorToOpenXml (op: ComparisonOperator) : ConditionalFormattingOperatorValues =
        match op with
        | Equal -> ConditionalFormattingOperatorValues.Equal
        | NotEqual -> ConditionalFormattingOperatorValues.NotEqual
        | GreaterThan -> ConditionalFormattingOperatorValues.GreaterThan
        | LessThan -> ConditionalFormattingOperatorValues.LessThan
        | GreaterThanOrEqual -> ConditionalFormattingOperatorValues.GreaterThanOrEqual
        | LessThanOrEqual -> ConditionalFormattingOperatorValues.LessThanOrEqual
        | Between -> ConditionalFormattingOperatorValues.Between
        | NotBetween -> ConditionalFormattingOperatorValues.NotBetween

    let private cfvo (valueType: ConditionalFormatValueObjectValues) (value: string option) : ConditionalFormatValueObject =
        let v = ConditionalFormatValueObject(Type = EnumValue<ConditionalFormatValueObjectValues>(valueType))
        value |> Option.iter (fun s -> v.Val <- StringValue(s))
        v

    /// Builds a `cfRule` element for one `ConditionalFormatRule`. `priority` becomes the
    /// rule's evaluation priority (lower numbers are evaluated first); `registry` is where
    /// `CellValueRule`/`FormulaRule`/`DuplicateValuesRule`/`UniqueValuesRule` intern their
    /// style into the stylesheet's `dxfs` collection - `ColorScale`/`DataBar` rules define
    /// their colors inline instead, so they never touch the registry.
    let private conditionalFormattingRuleElement
        (registry: StyleRegistry)
        (priority: int)
        (rule: ConditionalFormatRule)
        : ConditionalFormattingRule =
        let cfRule = ConditionalFormattingRule(Priority = Int32Value(priority))

        match rule with
        | CellValueRule(operator, formula1, formula2, style) ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.CellIs)
            cfRule.Operator <- EnumValue<ConditionalFormattingOperatorValues>(comparisonOperatorToOpenXml operator)
            cfRule.FormatId <- UInt32Value(registry.InternDxf style)
            cfRule.AppendChild(Spreadsheet.Formula(formula1)) |> ignore
            formula2 |> Option.iter (fun f2 -> cfRule.AppendChild(Spreadsheet.Formula(f2)) |> ignore)
        | FormulaRule(formula, style) ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.Expression)
            cfRule.FormatId <- UInt32Value(registry.InternDxf style)
            cfRule.AppendChild(Spreadsheet.Formula(formula)) |> ignore
        | ColorScale2(minColor, maxColor) ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.ColorScale)
            let cs = ColorScale()
            cs.AppendChild(cfvo ConditionalFormatValueObjectValues.Min None) |> ignore
            cs.AppendChild(cfvo ConditionalFormatValueObjectValues.Max None) |> ignore
            cs.AppendChild(ColorMapping.colorElement minColor) |> ignore
            cs.AppendChild(ColorMapping.colorElement maxColor) |> ignore
            cfRule.AppendChild(cs) |> ignore
        | ColorScale3(minColor, midColor, maxColor) ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.ColorScale)
            let cs = ColorScale()
            cs.AppendChild(cfvo ConditionalFormatValueObjectValues.Min None) |> ignore
            // Matches Excel's own default midpoint when you apply a 3-color scale from the UI.
            cs.AppendChild(cfvo ConditionalFormatValueObjectValues.Percentile (Some "50")) |> ignore
            cs.AppendChild(cfvo ConditionalFormatValueObjectValues.Max None) |> ignore
            cs.AppendChild(ColorMapping.colorElement minColor) |> ignore
            cs.AppendChild(ColorMapping.colorElement midColor) |> ignore
            cs.AppendChild(ColorMapping.colorElement maxColor) |> ignore
            cfRule.AppendChild(cs) |> ignore
        | DataBarRule color ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.DataBar)
            let db = DataBar()
            db.AppendChild(cfvo ConditionalFormatValueObjectValues.Min None) |> ignore
            db.AppendChild(cfvo ConditionalFormatValueObjectValues.Max None) |> ignore
            db.AppendChild(ColorMapping.colorElement color) |> ignore
            cfRule.AppendChild(db) |> ignore
        | DuplicateValuesRule style ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.DuplicateValues)
            cfRule.FormatId <- UInt32Value(registry.InternDxf style)
        | UniqueValuesRule style ->
            cfRule.Type <- EnumValue<ConditionalFormatValues>(ConditionalFormatValues.UniqueValues)
            cfRule.FormatId <- UInt32Value(registry.InternDxf style)

        cfRule

    /// Builds one `conditionalFormatting` container - one per `ConditionalFormatEntry`
    /// (i.e. one range, one rule), rather than trying to group several rules that
    /// happen to share a range into a single container. Simpler, and just as valid.
    let private conditionalFormattingElement (registry: StyleRegistry) (priority: int) (entry: ConditionalFormatEntry) : ConditionalFormatting =
        let cf = ConditionalFormatting()
        cf.SequenceOfReferences <- ListValue<StringValue>([ StringValue(rangeReference entry.TopLeft entry.BottomRight) ])
        cf.AppendChild(conditionalFormattingRuleElement registry priority entry.Rule) |> ignore
        cf

    let private dataValidationOperatorToOpenXml (op: ComparisonOperator) : DataValidationOperatorValues =
        match op with
        | Equal -> DataValidationOperatorValues.Equal
        | NotEqual -> DataValidationOperatorValues.NotEqual
        | GreaterThan -> DataValidationOperatorValues.GreaterThan
        | LessThan -> DataValidationOperatorValues.LessThan
        | GreaterThanOrEqual -> DataValidationOperatorValues.GreaterThanOrEqual
        | LessThanOrEqual -> DataValidationOperatorValues.LessThanOrEqual
        | Between -> DataValidationOperatorValues.Between
        | NotBetween -> DataValidationOperatorValues.NotBetween

    let private errorStyleToOpenXml (s: ErrorAlertStyle) : DataValidationErrorStyleValues =
        match s with
        | Stop -> DataValidationErrorStyleValues.Stop
        | Warning -> DataValidationErrorStyleValues.Warning
        | Information -> DataValidationErrorStyleValues.Information

    /// OOXML's inline list source is a double-quoted, comma-separated formula string
    /// (e.g. `"Yes,No,Maybe"`) - a literal `"` inside an item is escaped by doubling it,
    /// the same convention spreadsheet formulas use for embedded quotes.
    let private listFormula (items: string list) : string =
        items |> List.map (fun s -> s.Replace("\"", "\"\"")) |> String.concat "," |> sprintf "\"%s\""

    let private pageMarginsElement (m: PageMargins) : Spreadsheet.PageMargins =
        Spreadsheet.PageMargins(
            Left = DoubleValue(m.Left),
            Right = DoubleValue(m.Right),
            Top = DoubleValue(m.Top),
            Bottom = DoubleValue(m.Bottom),
            Header = DoubleValue(m.Header),
            Footer = DoubleValue(m.Footer)
        )

    /// Builds a `pageSetup` element. `Scale`/`FitToWidth`/`FitToHeight` are both always
    /// written when applicable - Excel only actually *uses* whichever one the sibling
    /// `sheetPr/pageSetUpPr/@fitToPage` flag selects (see `populate`, which sets that flag
    /// precisely when `Scaling` is `FitToPage`), but the unused attribute doesn't hurt and
    /// this way the file is self-describing either way.
    let private pageSetupElement (ps: PageSetup) : Spreadsheet.PageSetup =
        let setup =
            Spreadsheet.PageSetup(Orientation = EnumValue<OrientationValues>(OrientationMapping.toOpenXml ps.Orientation))

        ps.PaperSize |> Option.iter (fun p -> setup.PaperSize <- UInt32Value(PaperSizeMapping.toOpenXml p))

        match ps.Scaling with
        | Some(ScalePercent pct) -> setup.Scale <- UInt32Value(uint32 pct)
        | Some(FitToPage(width, height)) ->
            setup.FitToWidth <- UInt32Value(uint32 width)
            setup.FitToHeight <- UInt32Value(uint32 height)
        | None -> ()

        setup

    let private headerFooterElement (ps: PageSetup) : HeaderFooter option =
        if ps.Header.IsNone && ps.Footer.IsNone then
            None
        else
            let hf = HeaderFooter()
            ps.Header |> Option.iter (fun h -> hf.OddHeader <- OddHeader(h))
            ps.Footer |> Option.iter (fun f -> hf.OddFooter <- OddFooter(f))
            Some hf

    let private tableColumnElement (id: uint32) (col: TableColumn) : Spreadsheet.TableColumn =
        let tc = Spreadsheet.TableColumn(Id = UInt32Value(id), Name = StringValue(col.Name))
        col.CalculatedFormula |> Option.iter (fun f -> tc.CalculatedColumnFormula <- CalculatedColumnFormula(f))
        tc

    let private tableStyleInfoElement (style: TableStyle) : TableStyleInfo =
        let tsi =
            TableStyleInfo(
                ShowFirstColumn = BooleanValue(style.ShowFirstColumn),
                ShowLastColumn = BooleanValue(style.ShowLastColumn),
                ShowRowStripes = BooleanValue(style.ShowRowStripes),
                ShowColumnStripes = BooleanValue(style.ShowColumnStripes)
            )

        style.Name |> Option.iter (fun n -> tsi.Name <- StringValue(n))
        tsi

    /// Builds a table part's root `table` element. Raises if `Columns` doesn't match the
    /// range's width or contains duplicate names - genuine caller mistakes Excel itself
    /// would refuse to open cleanly, not something to paper over (same philosophy as
    /// `definedNameElement`'s sheet-name check). `Name` is written to both `name` and
    /// `displayName` - see `TableEntry`'s own doc comment.
    let private tableElement (tableId: uint32) (entry: TableEntry) : Spreadsheet.Table =
        let width = entry.BottomRight.Col - entry.TopLeft.Col + 1

        if entry.Columns.Length <> width then
            invalidArg
                (nameof entry)
                (sprintf "Table '%s' has %d column(s) but its range is %d column(s) wide" entry.Name entry.Columns.Length width)

        let names = entry.Columns |> List.map (fun c -> c.Name)

        if names |> List.distinct |> List.length <> names.Length then
            invalidArg (nameof entry) (sprintf "Table '%s' has duplicate column names" entry.Name)

        let table =
            Spreadsheet.Table(
                Id = UInt32Value(tableId),
                Name = StringValue(entry.Name),
                DisplayName = StringValue(entry.Name),
                Reference = StringValue(rangeReference entry.TopLeft entry.BottomRight),
                HeaderRowCount = UInt32Value(1u),
                // Core never writes a totals row - see MAPPING.md - so this is always
                // explicit false, matching what Excel itself writes for a table with none.
                TotalsRowShown = BooleanValue(false)
            )

        table.AutoFilter <- Spreadsheet.AutoFilter(Reference = StringValue(rangeReference entry.TopLeft entry.BottomRight))

        let tableColumns = TableColumns(Count = UInt32Value(uint32 entry.Columns.Length))
        entry.Columns |> List.iteri (fun i col -> tableColumns.AppendChild(tableColumnElement (uint32 (i + 1)) col) |> ignore)
        table.TableColumns <- tableColumns

        table.TableStyleInfo <- tableStyleInfoElement entry.Style

        table

    let private dataValidationElement (entry: DataValidationEntry) : Spreadsheet.DataValidation =
        let dv = Spreadsheet.DataValidation()
        dv.SequenceOfReferences <- ListValue<StringValue>([ StringValue(rangeReference entry.TopLeft entry.BottomRight) ])
        dv.AllowBlank <- BooleanValue(entry.Alert.AllowBlank)
        dv.ErrorStyle <- EnumValue<DataValidationErrorStyleValues>(errorStyleToOpenXml entry.Alert.ErrorStyle)

        if entry.Alert.ErrorTitle.IsSome || entry.Alert.ErrorMessage.IsSome then
            dv.ShowErrorMessage <- BooleanValue(true)
            entry.Alert.ErrorTitle |> Option.iter (fun t -> dv.ErrorTitle <- StringValue(t))
            entry.Alert.ErrorMessage |> Option.iter (fun m -> dv.Error <- StringValue(m))

        if entry.Alert.InputTitle.IsSome || entry.Alert.InputMessage.IsSome then
            dv.ShowInputMessage <- BooleanValue(true)
            entry.Alert.InputTitle |> Option.iter (fun t -> dv.PromptTitle <- StringValue(t))
            entry.Alert.InputMessage |> Option.iter (fun m -> dv.Prompt <- StringValue(m))

        let setOperatorAndFormulas operator formula1 formula2 =
            dv.Operator <- EnumValue<DataValidationOperatorValues>(dataValidationOperatorToOpenXml operator)
            dv.Formula1 <- Formula1(formula1: string)
            formula2 |> Option.iter (fun f2 -> dv.Formula2 <- Formula2(f2))

        // Deliberately never sets ShowDropDown: OOXML inverts its meaning for list
        // validations - true *hides* the in-cell dropdown arrow - so leaving it unset
        // (the default) is what gives the normal, expected visible dropdown.
        match entry.Kind with
        | ListValidation items ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.List)
            dv.Formula1 <- Formula1(listFormula items)
        | ListFromRangeValidation(topLeft, bottomRight) ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.List)
            dv.Formula1 <- Formula1(rangeReference topLeft bottomRight)
        | WholeNumberValidation(operator, formula1, formula2) ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.Whole)
            setOperatorAndFormulas operator formula1 formula2
        | DecimalValidation(operator, formula1, formula2) ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.Decimal)
            setOperatorAndFormulas operator formula1 formula2
        | TextLengthValidation(operator, formula1, formula2) ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.TextLength)
            setOperatorAndFormulas operator formula1 formula2
        | CustomValidation formula ->
            dv.Type <- EnumValue<DataValidationValues>(DataValidationValues.Custom)
            dv.Formula1 <- Formula1(formula)

        dv

    let private populate (document: SpreadsheetDocument) (wb: Workbook) =
        let workbookPart = document.AddWorkbookPart()
        workbookPart.Workbook <- Spreadsheet.Workbook()
        let registry = StyleRegistry()

        let sharedStrings = ResizeArray<string>()
        let sharedStringIndex = Dictionary<string, int>()

        let internString (s: string) =
            match sharedStringIndex.TryGetValue s with
            | true, idx -> idx
            | false, _ ->
                let idx = sharedStrings.Count
                sharedStrings.Add s
                sharedStringIndex.[s] <- idx
                idx

        let sheetsElement = Sheets()
        workbookPart.Workbook.AppendChild(sheetsElement) |> ignore

        let stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>()

        // A table's `id` must be unique across the whole workbook, not just its own sheet
        // (ECMA-376 Part 1 §18.5.1.2), so this counts up across the entire sheet loop
        // below rather than restarting per sheet.
        let mutable nextTableId = 1u

        wb.Sheets
        |> List.iteri (fun i worksheet ->
            let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
            let ws = Spreadsheet.Worksheet()

            // `sheetPr` must be the worksheet's first child if present at all (schema
            // order), so this has to run before every other element below. Only written
            // for `FitToPage` scaling: that's the one case with a flag to set
            // (`pageSetUpPr/@fitToPage`) - Excel's own default (scale-percent mode) needs
            // no `sheetPr` at all.
            match worksheet.PageSetup |> Option.bind (fun ps -> ps.Scaling) with
            | Some(FitToPage _) ->
                let sheetPr = SheetProperties(PageSetupProperties = PageSetupProperties(FitToPage = BooleanValue(true)))
                ws.AppendChild(sheetPr) |> ignore
            | _ -> ()

            match worksheet.FreezePane with
            | Some fp when fp.Rows > 0 || fp.Columns > 0 ->
                let pane =
                    Pane(
                        State = EnumValue<PaneStateValues>(PaneStateValues.Frozen),
                        TopLeftCell = StringValue(CellRef.toA1 (CellRef.create fp.Rows fp.Columns))
                    )

                if fp.Columns > 0 then
                    pane.HorizontalSplit <- DoubleValue(float fp.Columns)

                if fp.Rows > 0 then
                    pane.VerticalSplit <- DoubleValue(float fp.Rows)

                pane.ActivePane <-
                    EnumValue<PaneValues>(
                        if fp.Rows > 0 && fp.Columns > 0 then PaneValues.BottomRight
                        elif fp.Rows > 0 then PaneValues.BottomLeft
                        else PaneValues.TopRight
                    )

                let sheetView = SheetView(WorkbookViewId = UInt32Value(0u))
                sheetView.AppendChild(pane) |> ignore
                let sheetViews = SheetViews()
                sheetViews.AppendChild(sheetView) |> ignore
                ws.AppendChild(sheetViews) |> ignore
            | _ -> ()

            if not worksheet.ColumnProps.IsEmpty then
                let columns = Columns()

                worksheet.ColumnProps
                |> Map.toList
                |> List.sortBy fst
                |> List.iter (fun (colIdx, props) ->
                    match props.Width with
                    | Some w ->
                        columns.AppendChild(
                            Column(
                                Min = UInt32Value(uint32 (colIdx + 1)),
                                Max = UInt32Value(uint32 (colIdx + 1)),
                                Width = DoubleValue(w),
                                CustomWidth = BooleanValue(true)
                            )
                        )
                        |> ignore
                    | None -> ())

                if columns.HasChildren then
                    ws.AppendChild(columns) |> ignore

            let sheetData = SheetData()

            let cellsByRow =
                worksheet.Cells
                |> List.groupBy (fun c -> c.Ref.Row)
                |> Map.ofList

            let allRowIndices =
                Set.union
                    (worksheet.Cells |> List.map (fun c -> c.Ref.Row) |> Set.ofList)
                    (worksheet.RowProps |> Map.toList |> List.map fst |> Set.ofList)
                |> Set.toList
                |> List.sort

            for rowIdx in allRowIndices do
                let row = Spreadsheet.Row(RowIndex = UInt32Value(uint32 (rowIdx + 1)))

                worksheet.RowProps
                |> Map.tryFind rowIdx
                |> Option.bind (fun rp -> rp.Height)
                |> Option.iter (fun h ->
                    row.Height <- DoubleValue(h)
                    row.CustomHeight <- BooleanValue(true))

                let cellsInRow =
                    cellsByRow
                    |> Map.tryFind rowIdx
                    |> Option.defaultValue []
                    |> List.sortBy (fun c -> c.Ref.Col)

                for cell in cellsInRow do
                    let baseStyle = defaultArg cell.Style CellStyle.Default

                    // A Date cell with no explicit number format needs one, or Excel will
                    // display the raw OLE Automation serial number instead of a date.
                    let effectiveStyle =
                        match cell.Value, baseStyle.NumberFormat with
                        | Date d, None ->
                            let hasTime = d.TimeOfDay <> TimeSpan.Zero
                            { baseStyle with
                                NumberFormat = Some(if hasTime then DateAndTime else ShortDate) }
                        | _ -> baseStyle

                    let styleIdx = registry.GetCellFormatIndex(Some effectiveStyle)
                    let c = Spreadsheet.Cell(CellReference = StringValue(CellRef.toA1 cell.Ref))

                    if styleIdx <> 0u then
                        c.StyleIndex <- UInt32Value(styleIdx)

                    match cell.Value with
                    | Empty -> ()
                    | Text s ->
                        let idx = internString s
                        c.DataType <- EnumValue<CellValues>(CellValues.SharedString)
                        c.CellValue <- Spreadsheet.CellValue(string idx)
                    | Number n -> c.CellValue <- Spreadsheet.CellValue(n.ToString(inv))
                    | Boolean b ->
                        c.DataType <- EnumValue<CellValues>(CellValues.Boolean)
                        c.CellValue <- Spreadsheet.CellValue(if b then "1" else "0")
                    | Date d -> c.CellValue <- Spreadsheet.CellValue(d.ToOADate().ToString(inv))
                    | Formula(expr, cached) ->
                        c.CellFormula <- CellFormula(Text = expr)
                        cached |> Option.iter (fun v -> c.CellValue <- Spreadsheet.CellValue(v.ToString(inv)))

                    row.AppendChild(c) |> ignore

                sheetData.AppendChild(row) |> ignore

            ws.AppendChild(sheetData) |> ignore

            worksheet.Protection
            |> Option.iter (fun sp -> ws.AppendChild(sheetProtectionElement sp) |> ignore)

            worksheet.AutoFilter
            |> Option.iter (fun range ->
                ws.AppendChild(Spreadsheet.AutoFilter(Reference = StringValue(rangeReference range.TopLeft range.BottomRight)))
                |> ignore)

            if not worksheet.MergedRanges.IsEmpty then
                let mergeCells = MergeCells()

                worksheet.MergedRanges
                |> List.iter (fun m ->
                    mergeCells.AppendChild(
                        MergeCell(
                            Reference = StringValue(sprintf "%s:%s" (CellRef.toA1 m.TopLeft) (CellRef.toA1 m.BottomRight))
                        )
                    )
                    |> ignore)

                ws.AppendChild(mergeCells) |> ignore

            worksheet.ConditionalFormats
            |> List.iteri (fun idx entry ->
                ws.AppendChild(conditionalFormattingElement registry (idx + 1) entry) |> ignore)

            if not worksheet.DataValidations.IsEmpty then
                let dvs = DataValidations(Count = UInt32Value(uint32 worksheet.DataValidations.Length))
                worksheet.DataValidations |> List.iter (fun entry -> dvs.AppendChild(dataValidationElement entry) |> ignore)
                ws.AppendChild(dvs) |> ignore

            if not worksheet.Hyperlinks.IsEmpty then
                let hyperlinks = Hyperlinks()

                worksheet.Hyperlinks
                |> List.iter (fun entry -> hyperlinks.AppendChild(hyperlinkElement worksheetPart entry) |> ignore)

                ws.AppendChild(hyperlinks) |> ignore

            worksheet.PageSetup
            |> Option.iter (fun ps ->
                ws.AppendChild(pageMarginsElement ps.Margins) |> ignore
                ws.AppendChild(pageSetupElement ps) |> ignore
                headerFooterElement ps |> Option.iter (fun hf -> ws.AppendChild(hf) |> ignore))

            if not worksheet.Comments.IsEmpty then
                let commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>()
                commentsPart.Comments <- commentsRoot worksheet.Comments
                commentsPart.Comments.Save()

                let vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>()

                use vmlStream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)
                use writer = new StreamWriter(vmlStream)
                writer.Write(vmlDrawingContent worksheet.Comments)
                writer.Flush()

                ws.AppendChild(LegacyDrawing(Id = StringValue(worksheetPart.GetIdOfPart(vmlPart)))) |> ignore

            if not worksheet.Tables.IsEmpty then
                let tableParts = TableParts(Count = UInt32Value(uint32 worksheet.Tables.Length))

                worksheet.Tables
                |> List.iter (fun entry ->
                    let tableDefPart = worksheetPart.AddNewPart<TableDefinitionPart>()
                    tableDefPart.Table <- tableElement nextTableId entry
                    tableDefPart.Table.Save()
                    nextTableId <- nextTableId + 1u
                    tableParts.AppendChild(TablePart(Id = StringValue(worksheetPart.GetIdOfPart(tableDefPart)))) |> ignore)

                ws.AppendChild(tableParts) |> ignore

            worksheetPart.Worksheet <- ws
            worksheetPart.Worksheet.Save()

            let sheetId = uint32 (i + 1)
            let relId = workbookPart.GetIdOfPart(worksheetPart)

            sheetsElement.AppendChild(
                Sheet(Name = StringValue(worksheet.Name), SheetId = UInt32Value(sheetId), Id = StringValue(relId))
            )
            |> ignore)

        if not wb.DefinedNames.IsEmpty then
            let sheetIndex = wb.Sheets |> List.mapi (fun i s -> s.Name, i) |> Map.ofList
            let definedNames = DefinedNames()

            wb.DefinedNames
            |> List.iter (fun entry -> definedNames.AppendChild(definedNameElement sheetIndex entry) |> ignore)

            workbookPart.Workbook.AppendChild(definedNames) |> ignore

        let sstPart = workbookPart.AddNewPart<SharedStringTablePart>()
        let sst = SharedStringTable()
        sharedStrings
        |> Seq.iter (fun s ->
            let item = SharedStringItem()
            item.AppendChild(Spreadsheet.Text(s)) |> ignore
            sst.AppendChild(item) |> ignore)
        sstPart.SharedStringTable <- sst
        sstPart.SharedStringTable.Save()

        stylesPart.Stylesheet <- registry.BuildStylesheet()
        stylesPart.Stylesheet.Save()

        workbookPart.Workbook.Save()

    let saveToStream (wb: Workbook) (stream: Stream) : unit =
        use document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
        populate document wb

    let saveToFile (wb: Workbook) (path: string) : unit =
        use document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
        populate document wb
