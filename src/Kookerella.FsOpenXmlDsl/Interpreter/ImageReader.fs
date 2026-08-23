namespace Kookerella.FsOpenXmlDsl.Interpreter

open System.IO
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsOpenXmlDsl

/// Parses the images anchored on one worksheet back into `ImageEntry` values - the read
/// side of `ImageWriter`. See `ChartReader`'s own doc comment for why this is a separate
/// file and why none of `DocumentFormat.OpenXml.Drawing`/`.Drawing.Spreadsheet` are
/// `open`ed.
module internal ImageReader =

    type XTwoCellAnchor = DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor
    type XPicture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture

    /// Only the four formats `ImageWriter` ever writes are recognized - a foreign file's
    /// image in some other format (TIFF, SVG, EMF/WMF, ...) is dropped rather than
    /// guessed at, same "best-effort" philosophy as the rest of this module.
    let private formatOfContentType (contentType: string) : ImageFormat option =
        match contentType with
        | "image/png" -> Some Png
        | "image/jpeg" -> Some Jpeg
        | "image/gif" -> Some Gif
        | "image/bmp" -> Some Bmp
        | _ -> None

    /// Tries to interpret one anchor as an image - `None` if it doesn't contain a
    /// `pic`/image relationship at all (e.g. it's a chart instead - see
    /// `ChartReader.tryReadChart`) or its image isn't one of the four formats
    /// `ImageWriter` ever writes.
    let tryReadImage (drawingsPart: DrawingsPart) (anchor: XTwoCellAnchor) (topLeftAnchor: CellRef) (bottomRightAnchor: CellRef) : ImageEntry option =
        anchor.Elements<XPicture>()
        |> Seq.tryHead
        |> Option.bind (fun pic -> Option.ofObj pic.BlipFill)
        |> Option.bind (fun bf -> Option.ofObj bf.Blip)
        |> Option.bind (fun blip -> Option.ofObj blip.Embed)
        |> Option.bind (fun relIdVal ->
            match drawingsPart.GetPartById(relIdVal.Value) with
            | :? ImagePart as imagePart -> Some imagePart
            | _ -> None)
        |> Option.bind (fun imagePart ->
            formatOfContentType imagePart.ContentType
            |> Option.map (fun format ->
                use stream = imagePart.GetStream(FileMode.Open, FileAccess.Read)
                use buffer = new MemoryStream()
                stream.CopyTo(buffer)

                { Data = buffer.ToArray()
                  Format = format
                  TopLeftAnchor = topLeftAnchor
                  BottomRightAnchor = bottomRightAnchor }))
