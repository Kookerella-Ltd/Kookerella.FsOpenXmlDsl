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
