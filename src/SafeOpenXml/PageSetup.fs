namespace SafeOpenXml

/// Page orientation for printing.
type PageOrientation =
    | Portrait
    | Landscape

/// A small named set of common paper sizes (OOXML's `ST_PaperSize` numeric codes, ECMA-376
/// Part 1 §18.18.50) plus `OtherPaperSize` as an escape hatch for any of the several dozen
/// remaining codes that aren't worth naming individually - same shape as
/// `BorderLineStyle.Other`, just under a distinct name: `BorderLineStyle`/`NumberFormat`
/// both already have their own `Other`/`Custom` escape-hatch case, and `open SafeOpenXml`
/// bringing all three types into scope at once means a bare `Other` would silently
/// resolve to whichever type was declared last, not necessarily the one being matched on.
type PaperSize =
    | Letter
    | Legal
    | Tabloid
    | A3
    | A4
    | OtherPaperSize of code: int

/// How the printed sheet is scaled onto its pages - Excel's print dialog toggles between
/// these two mutually exclusive modes ("Adjust to" vs. "Fit to"). `FitToPage`'s
/// `width`/`height` are page counts; `0` means "as many as needed" in that dimension,
/// matching Excel's own convention (e.g. "fit to 1 page wide, any number tall" is
/// `FitToPage(1, 0)`).
type PrintScaling =
    | ScalePercent of percent: int
    | FitToPage of width: int * height: int

/// Margins in inches, matching OOXML's `pageMargins` directly - there is no cleaner unit
/// to convert to/from.
type PageMargins =
    { Left: float
      Right: float
      Top: float
      Bottom: float
      Header: float
      Footer: float }

    /// Excel's own built-in margins - what a fresh worksheet prints with even without an
    /// explicit `pageMargins` element at all.
    static member Default =
        { Left = 0.7
          Right = 0.7
          Top = 0.75
          Bottom = 0.75
          Header = 0.3
          Footer = 0.3 }

/// Print settings for a worksheet - orientation, paper size, scaling, margins, print area,
/// and a header/footer (including its first-page/even-page variants). `Header`/`Footer`
/// (and their `Even`/`First` counterparts) are raw OOXML header/footer text: Excel's own
/// `&L`/`&C`/`&R` (left/center/right section) and `&P`/`&N`/`&D`/`&T`/`&F`/`&A` (page
/// number/total pages/date/time/filename/sheet name) codes embedded directly in one
/// string, the same convention as OOXML's own `oddHeader`/`oddFooter`/etc. `Header`/
/// `Footer` are shown on every page unless overridden: `EvenHeader`/`EvenFooter` apply to
/// even pages when set (Excel falls back to `Header`/`Footer` for even pages otherwise),
/// and `FirstHeader`/`FirstFooter` apply only to page 1.
///
/// `PrintArea` is a list of ranges (Excel supports printing several disjoint
/// rectangles as one print area) - empty means "no print area set", i.e. Excel prints the
/// whole used range. Unlike every other field here, which maps onto worksheet-element
/// attributes directly, `PrintArea` is actually a reserved, hidden, sheet-scoped defined
/// name (`_xlnm.Print_Area`) under the hood - `Writer`/`Reader` translate transparently,
/// the same way `DefinedNameScope.SheetScope` translates to/from OOXML's `localSheetId`,
/// so this is the one field on `PageSetup` that doesn't correspond to a `pageSetup`/
/// `pageMargins`/`headerFooter` attribute at all. Setting only `PrintArea` (leaving every
/// other field at its default) still causes `pageMargins`/`pageSetup` to be written using
/// their own defaults, same as setting any other field here.
type PageSetup =
    { Orientation: PageOrientation
      PaperSize: PaperSize option
      Scaling: PrintScaling option
      Margins: PageMargins
      PrintArea: (CellRef * CellRef) list
      Header: string option
      Footer: string option
      EvenHeader: string option
      EvenFooter: string option
      FirstHeader: string option
      FirstFooter: string option }

    static member Default =
        { Orientation = Portrait
          PaperSize = None
          Scaling = None
          Margins = PageMargins.Default
          PrintArea = []
          Header = None
          Footer = None
          EvenHeader = None
          EvenFooter = None
          FirstHeader = None
          FirstFooter = None }
