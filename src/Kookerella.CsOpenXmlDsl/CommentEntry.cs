namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A classic cell comment - what the OOXML spec calls a <c>comment</c> and current Excel's
/// UI now calls a "Note" (Excel's newer @mention/reply "Comments" are a different, separate
/// part format not modeled here). <see cref="Author"/> may be empty, matching how Excel
/// itself allows an unnamed comment author. Mirrors the F# core's <c>CommentEntry</c>.
/// </summary>
public sealed record CommentEntry(CellPosition Cell, string Text, string Author = "")
{
    public static CommentEntry Of(string cellA1, string text, string author = "") =>
        new(CellPosition.FromA1(cellA1), text, author);
}
