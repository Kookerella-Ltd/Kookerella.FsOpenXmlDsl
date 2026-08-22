namespace SafeOpenXml

/// The visual shape of a sparkline group - Excel's own three kinds. `WinLoss` is what
/// Excel's UI calls "Win/Loss" (OOXML names the same enum value `stacked`).
type SparklineType =
    | Line
    | Column
    | WinLoss

/// One sparkline's target cell and the data range it summarizes. `DataTopLeft`/
/// `DataBottomRight` are typically a single row or column (the common case: one
/// sparkline per row of data), but any rectangular range is accepted, same as everywhere
/// else ranges appear in this DSL.
type SparklineCell =
    { Cell: CellRef
      DataTopLeft: CellRef
      DataBottomRight: CellRef }

/// Shared visual settings for a group of sparklines - Excel groups sparklines so several
/// can be styled together at once (e.g. a filled-down column of row-summary sparklines).
/// Covers the commonly used subset: `Color` is the one color Core models (the main
/// sparkline color itself; Excel's separate negative-point/axis/marker colors default to
/// its own automatic choices) and `ShowHigh`/`ShowLow`/`ShowFirst`/`ShowLast`/
/// `ShowNegative` are the "highlight these points" toggles from Excel's Sparkline Design
/// ribbon. See MAPPING.md for the richer options not modeled (axis settings, date axis,
/// per-role marker/negative colors, hidden/empty-cell handling).
type SparklineStyle =
    { Type: SparklineType
      Color: Color option
      /// Points; only meaningful for `Line`.
      LineWeight: float option
      /// Only meaningful for `Line`.
      ShowMarkers: bool
      ShowHigh: bool
      ShowLow: bool
      ShowFirst: bool
      ShowLast: bool
      ShowNegative: bool }

    static member Default =
        { Type = Line
          Color = None
          LineWeight = None
          ShowMarkers = false
          ShowHigh = false
          ShowLow = false
          ShowFirst = false
          ShowLast = false
          ShowNegative = false }

/// A group of sparklines sharing one `SparklineStyle`, as stored on `Worksheet` - a sheet
/// can have several independently-styled groups (e.g. one `Line` group and one `Column`
/// group in different columns).
type SparklineGroupEntry =
    { Style: SparklineStyle
      Sparklines: SparklineCell list }
