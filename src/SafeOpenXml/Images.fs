namespace SafeOpenXml

/// The image's file format - determines the OOXML content type its `ImagePart` is
/// registered with. Covers the four formats every Excel version has always supported
/// natively; see MAPPING.md for less common ones (TIFF, SVG, EMF/WMF, ...) that aren't
/// modeled.
type ImageFormat =
    | Png
    | Jpeg
    | Gif
    | Bmp

/// A raster image anchored over a range of cells on a worksheet, as stored on `Worksheet`
/// (a sheet can have several). `Data` is the image file's own raw bytes exactly as they'd
/// sit on disk - this DSL does no image decoding/encoding of its own, it only embeds
/// whatever bytes you give it (and hands back exactly those bytes on read). `TopLeftAnchor`/
/// `BottomRightAnchor` size and position the image by spanning exactly that range of cells
/// (a "move and size with cells" anchor, snapped to cell boundaries), the same convention
/// `ChartEntry`/`TableEntry`/`MergedRange` already use, rather than pixel-precise floating
/// position - see MAPPING.md for what isn't modeled (free-floating position, rotation,
/// cropping, alt text).
type ImageEntry =
    { Data: byte[]
      Format: ImageFormat
      TopLeftAnchor: CellRef
      BottomRightAnchor: CellRef }
