namespace Kookerella.FsOpenXmlDsl

/// Where a hyperlink points. `ExternalHyperlink` covers ordinary URLs and `mailto:`
/// addresses alike - OOXML treats both as an external relationship, just with a
/// different URI scheme, so there's no separate email case. `InternalHyperlink` is a
/// same-workbook reference such as `"Sheet2!A1"` or a defined name, written to the
/// `location` attribute directly - no relationship needed for that one.
type HyperlinkTarget =
    | ExternalHyperlink of url: string
    | InternalHyperlink of location: string

/// A hyperlink applied over a range (a single cell is the degenerate case where
/// `TopLeft = BottomRight`), as it's stored on `Worksheet`.
///
/// `Display` is OOXML's fallback label - a handful of older Excel versions and some
/// interop tools show it instead of the cell's own text; modern Excel always shows the
/// cell's actual content and ignores it. Rarely worth setting, but preserved for
/// round-trip fidelity when reading a file that has one.
type HyperlinkEntry =
    { TopLeft: CellRef
      BottomRight: CellRef
      Target: HyperlinkTarget
      Tooltip: string option
      Display: string option }
