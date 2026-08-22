namespace SafeOpenXml.Interpreter

open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open SafeOpenXml

/// Builds the DrawingML parts and relationships for the images anchored on one worksheet
/// - the write side of `ImageReader`. Shares one `DrawingsPart` per worksheet with
/// `ChartWriter` (see `DrawingWriter`, which owns that shared lifecycle) rather than
/// creating its own, since a worksheet only gets one `<drawing>` relationship regardless
/// of how many charts/images it has.
///
/// See `ChartWriter`'s own doc comment for why none of `DocumentFormat.OpenXml.Drawing`/
/// `.Drawing.Spreadsheet` are `open`ed here either.
module internal ImageWriter =

    type XPicture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture
    type XNonVisualPictureProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties
    type XNonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties
    type XNonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
    type XBlipFill = DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill
    type XShapeProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties
    type XTwoCellAnchor = DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor
    type XFromMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker
    type XToMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker
    type XColumnId = DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId
    type XRowId = DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId
    type XColumnOffset = DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset
    type XRowOffset = DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset
    type XClientData = DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData

    type ABlip = DocumentFormat.OpenXml.Drawing.Blip
    type AStretch = DocumentFormat.OpenXml.Drawing.Stretch
    type AFillRectangle = DocumentFormat.OpenXml.Drawing.FillRectangle
    type ATransform2D = DocumentFormat.OpenXml.Drawing.Transform2D
    type AOffset = DocumentFormat.OpenXml.Drawing.Offset
    type AExtents = DocumentFormat.OpenXml.Drawing.Extents
    type APresetGeometry = DocumentFormat.OpenXml.Drawing.PresetGeometry
    type AAdjustValueList = DocumentFormat.OpenXml.Drawing.AdjustValueList
    type AShapeTypeValues = DocumentFormat.OpenXml.Drawing.ShapeTypeValues

    let private imagePartType (format: ImageFormat) : PartTypeInfo =
        match format with
        | Png -> ImagePartType.Png
        | Jpeg -> ImagePartType.Jpeg
        | Gif -> ImagePartType.Gif
        | Bmp -> ImagePartType.Bmp

    let private fromMarker (cell: CellRef) : XFromMarker =
        let m = XFromMarker()
        m.ColumnId <- XColumnId(string cell.Col)
        m.ColumnOffset <- XColumnOffset("0")
        m.RowId <- XRowId(string cell.Row)
        m.RowOffset <- XRowOffset("0")
        m

    /// Same "one past the cell" convention as `ChartWriter.toMarker`.
    let private toMarker (cell: CellRef) : XToMarker =
        let m = XToMarker()
        m.ColumnId <- XColumnId(string (cell.Col + 1))
        m.ColumnOffset <- XColumnOffset("0")
        m.RowId <- XRowId(string (cell.Row + 1))
        m.RowOffset <- XRowOffset("0")
        m

    let private pictureElement (imageId: uint32) (relId: string) : XPicture =
        let nvProps = XNonVisualPictureProperties()
        nvProps.NonVisualDrawingProperties <- XNonVisualDrawingProperties(Id = UInt32Value(imageId), Name = StringValue(sprintf "Picture %d" imageId))
        nvProps.NonVisualPictureDrawingProperties <- XNonVisualPictureDrawingProperties()

        let blipFill = XBlipFill()
        blipFill.Blip <- ABlip(Embed = StringValue(relId))
        blipFill.AppendChild(AStretch(FillRectangle = AFillRectangle())) |> ignore

        let shapeProps = XShapeProperties()

        shapeProps.Transform2D <-
            ATransform2D(Offset = AOffset(X = Int64Value(0L), Y = Int64Value(0L)), Extents = AExtents(Cx = Int64Value(0L), Cy = Int64Value(0L)))

        shapeProps.AppendChild(APresetGeometry(Preset = EnumValue<AShapeTypeValues>(AShapeTypeValues.Rectangle), AdjustValueList = AAdjustValueList()))
        |> ignore

        let picture = XPicture()
        picture.NonVisualPictureProperties <- nvProps
        picture.BlipFill <- blipFill
        picture.ShapeProperties <- shapeProps
        picture

    let private twoCellAnchorElement (entry: ImageEntry) (imageId: uint32) (relId: string) : XTwoCellAnchor =
        let anchor = XTwoCellAnchor()
        anchor.FromMarker <- fromMarker entry.TopLeftAnchor
        anchor.ToMarker <- toMarker entry.BottomRightAnchor
        anchor.AppendChild(pictureElement imageId relId) |> ignore
        anchor.AppendChild(XClientData()) |> ignore
        anchor

    /// Adds one `ImagePart` per image to `drawingsPart` (already created by the caller -
    /// see `chartAnchors`'s own doc comment) and returns the anchor element for each,
    /// starting numbering at `startId`.
    let imageAnchors (drawingsPart: DrawingsPart) (startId: uint32) (images: ImageEntry list) : OpenXmlElement list =
        images
        |> List.mapi (fun i entry ->
            let imagePart = drawingsPart.AddImagePart(imagePartType entry.Format)

            use stream = new MemoryStream(entry.Data)
            imagePart.FeedData(stream)

            let relId = drawingsPart.GetIdOfPart(imagePart)
            let imageId = startId + uint32 i
            twoCellAnchorElement entry imageId relId :> OpenXmlElement)
