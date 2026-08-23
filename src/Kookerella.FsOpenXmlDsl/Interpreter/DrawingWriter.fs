namespace Kookerella.FsOpenXmlDsl.Interpreter

open DocumentFormat.OpenXml.Packaging
open Kookerella.FsOpenXmlDsl

/// Owns the one `DrawingsPart` a worksheet gets when it has charts and/or images anchored
/// on it - both `ChartWriter.chartAnchors` and `ImageWriter.imageAnchors` just add their
/// own parts into a `DrawingsPart` this module creates and hands them, and return their
/// anchor elements for this module to collect into one shared `WorksheetDrawing` root
/// (a worksheet has exactly one `<drawing>` relationship regardless of how many charts/
/// images it has, so the two features can't each manage their own part).
module internal DrawingWriter =

    type private XWorksheetDrawing = DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing

    /// Builds the shared drawing canvas for one worksheet's charts and images, and
    /// returns the relationship id the worksheet's own `<drawing>` element should
    /// reference - `None` if the sheet has neither, so the caller (`Writer.populate`)
    /// knows not to add that element at all.
    let addDrawing (worksheetPart: WorksheetPart) (sheetName: string) (charts: ChartEntry list) (images: ImageEntry list) : string option =
        if charts.IsEmpty && images.IsEmpty then
            None
        else
            let drawingsPart = worksheetPart.AddNewPart<DrawingsPart>()
            let worksheetDrawing = XWorksheetDrawing()

            ChartWriter.chartAnchors drawingsPart sheetName 1u charts
            |> List.iter (fun a -> worksheetDrawing.AppendChild(a) |> ignore)

            ImageWriter.imageAnchors drawingsPart (1u + uint32 charts.Length) images
            |> List.iter (fun a -> worksheetDrawing.AppendChild(a) |> ignore)

            drawingsPart.WorksheetDrawing <- worksheetDrawing
            drawingsPart.WorksheetDrawing.Save()
            Some(worksheetPart.GetIdOfPart(drawingsPart))
