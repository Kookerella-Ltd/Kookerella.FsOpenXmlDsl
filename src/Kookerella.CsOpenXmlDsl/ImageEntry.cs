namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A raster image anchored over a range of cells on a worksheet (a sheet can have several).
/// <see cref="Data"/> is the image file's own raw bytes exactly as they'd sit on disk - this
/// wrapper does no image decoding/encoding of its own, it only embeds whatever bytes you give
/// it and hands back exactly those bytes on read. <see cref="TopLeftAnchor"/>/<see
/// cref="BottomRightAnchor"/> size and position the image by spanning exactly that range of
/// cells - the same "move and size with cells" anchor already used by
/// <see cref="ChartEntry"/>/<see cref="TableEntry"/>/<see cref="MergedRange"/> - rather than
/// pixel-precise floating position. Mirrors the F# core's <c>ImageEntry</c>.
/// <para>
/// This is the one place on an otherwise-pure type that can't fully guarantee immutability:
/// the constructor defensively copies <c>data</c>, so mutating the caller's original array
/// afterwards can't affect this entry - but the array returned from <see cref="Data"/> is
/// this record's own, so callers must not write to it.
/// </para>
/// </summary>
public sealed record ImageEntry
{
    private readonly byte[] _data;

    public byte[] Data => _data;
    public ImageFormat Format { get; }
    public CellPosition TopLeftAnchor { get; }
    public CellPosition BottomRightAnchor { get; }

    public ImageEntry(byte[] data, ImageFormat format, CellPosition topLeftAnchor, CellPosition bottomRightAnchor)
    {
        _data = data.ToArray();
        Format = format;
        TopLeftAnchor = topLeftAnchor;
        BottomRightAnchor = bottomRightAnchor;
    }

    public static ImageEntry Of(byte[] data, ImageFormat format, string topLeftAnchorA1, string bottomRightAnchorA1) =>
        new(data, format, CellPosition.FromA1(topLeftAnchorA1), CellPosition.FromA1(bottomRightAnchorA1));
}
