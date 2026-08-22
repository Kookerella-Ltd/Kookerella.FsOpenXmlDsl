namespace SafeOpenXml.Interpreter

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// Reverses `DrawingWriter.addDrawing`: walks a worksheet's `<drawing>` relationship to
/// its `DrawingsPart`, parses each anchor's cell range once, then tries it as a chart and
/// as an image - each anchor is really only ever one or the other, so this doesn't cost
/// anything to try both rather than have `ChartReader`/`ImageReader` duplicate the same
/// relationship-walking and marker-parsing code twice.
module internal DrawingReader =

    type private XTwoCellAnchor = DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor

    let private cellRefOfMarker (columnIdText: string) (rowIdText: string) : CellRef =
        CellRef.create (int rowIdText) (int columnIdText)

    /// A chart/image this can't make sense of (no categories, an unrecognized image
    /// format, ...) is dropped rather than failing the whole load, same "best-effort"
    /// philosophy as the rest of this module.
    let readChartsAndImages (worksheetPart: WorksheetPart) (ws: Spreadsheet.Worksheet) : ChartEntry list * ImageEntry list =
        let anchorsWithCells =
            ws.Elements<Drawing>()
            |> Seq.tryHead
            |> Option.bind (fun d -> Option.ofObj d.Id)
            |> Option.bind (fun relId ->
                match worksheetPart.GetPartById(relId.Value) with
                | :? DrawingsPart as dp -> Some dp
                | _ -> None)
            |> Option.bind (fun drawingsPart -> Option.ofObj drawingsPart.WorksheetDrawing |> Option.map (fun wd -> drawingsPart, wd))
            |> Option.map (fun (drawingsPart, worksheetDrawing) ->
                worksheetDrawing.Elements<XTwoCellAnchor>()
                |> Seq.choose (fun anchor ->
                    match Option.ofObj anchor.FromMarker, Option.ofObj anchor.ToMarker with
                    | Some fromMarker, Some toMarker ->
                        let fromCell = cellRefOfMarker fromMarker.ColumnId.Text fromMarker.RowId.Text
                        let toCellExclusive = cellRefOfMarker toMarker.ColumnId.Text toMarker.RowId.Text
                        let toCell = CellRef.create (toCellExclusive.Row - 1) (toCellExclusive.Col - 1)
                        Some(drawingsPart, anchor, fromCell, toCell)
                    | _ -> None)
                |> List.ofSeq)
            |> Option.defaultValue []

        let charts =
            anchorsWithCells
            |> List.choose (fun (dp, anchor, fromCell, toCell) -> ChartReader.tryReadChart dp anchor fromCell toCell)

        let images =
            anchorsWithCells
            |> List.choose (fun (dp, anchor, fromCell, toCell) -> ImageReader.tryReadImage dp anchor fromCell toCell)

        (charts, images)
