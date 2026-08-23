namespace Kookerella.FsOpenXmlDsl

/// A classic cell comment - what the OOXML spec calls a `comment` and current Excel's UI
/// now calls a "Note" (Excel's newer @mention/reply "Comments" are a different, separate
/// part format not modeled here - see MAPPING.md). `Author` may be empty, matching how
/// Excel itself allows an unnamed comment author.
type CommentEntry =
    { Cell: CellRef
      Author: string
      Text: string }
