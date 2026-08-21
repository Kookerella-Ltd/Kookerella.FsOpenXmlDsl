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

        wb.Sheets
        |> List.iteri (fun i worksheet ->
            let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
            let ws = Spreadsheet.Worksheet()

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

            worksheetPart.Worksheet <- ws
            worksheetPart.Worksheet.Save()

            let sheetId = uint32 (i + 1)
            let relId = workbookPart.GetIdOfPart(worksheetPart)

            sheetsElement.AppendChild(
                Sheet(Name = StringValue(worksheet.Name), SheetId = UInt32Value(sheetId), Id = StringValue(relId))
            )
            |> ignore)

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
