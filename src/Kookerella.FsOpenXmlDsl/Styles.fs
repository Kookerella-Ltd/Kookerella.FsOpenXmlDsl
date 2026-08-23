namespace Kookerella.FsOpenXmlDsl

/// Cell formatting. Structural equality on these types is load-bearing: the interpreter
/// interns identical `CellStyle` values into shared stylesheet entries (fonts/fills/borders/
/// number formats), the same way Excel itself deduplicates styles.
[<AutoOpen>]
module Styles =

    /// A cell color. `Rgb` is the only variant the DSL can meaningfully *author* -
    /// `Indexed` and `Theme` exist so that reading a real workbook back into the DSL is
    /// lossless even when the file uses the legacy indexed palette or a theme color/tint.
    /// Kookerella.FsOpenXmlDsl does not resolve `Theme`/`Indexed` to actual RGB values (that requires
    /// parsing the workbook's theme part, which Core does not model) - round-tripping such
    /// a file preserves the reference faithfully, but constructing one from scratch is
    /// only useful if you already know the target theme's palette.
    type Color =
        | Rgb of red: byte * green: byte * blue: byte
        | Indexed of paletteIndex: int
        | Theme of themeIndex: int * tint: float option

    module Color =
        let black = Rgb(0uy, 0uy, 0uy)
        let white = Rgb(255uy, 255uy, 255uy)
        let red = Rgb(255uy, 0uy, 0uy)
        let green = Rgb(0uy, 128uy, 0uy)
        let blue = Rgb(0uy, 0uy, 255uy)
        let yellow = Rgb(255uy, 255uy, 0uy)

    type FontStyle =
        { Name: string option
          /// Points.
          Size: float option
          Bold: bool
          Italic: bool
          Underline: bool
          Strikethrough: bool
          Color: Color option }

        static member Default =
            { Name = None
              Size = None
              Bold = false
              Italic = false
              Underline = false
              Strikethrough = false
              Color = None }

    /// Core only models a solid fill. Pattern fills (stripes, checkerboards, gradients)
    /// are a documented gap - see MAPPING.md.
    type FillStyle =
        { Color: Color }

    /// OOXML defines ~13 border line styles; Core covers the common ones. `Other` preserves
    /// the raw OOXML style name (e.g. "mediumDashed", "slantDashDot") so reading and
    /// re-writing an existing file round-trips even for styles Core doesn't author itself.
    type BorderLineStyle =
        | Thin
        | Medium
        | Thick
        | Dashed
        | Dotted
        | Double
        | Hair
        | Other of string

    type BorderSide =
        { Style: BorderLineStyle
          Color: Color option }

    type BorderStyle =
        { Left: BorderSide option
          Right: BorderSide option
          Top: BorderSide option
          Bottom: BorderSide option }

        static member None =
            { Left = None; Right = None; Top = None; Bottom = None }

    type HorizontalAlignment =
        | GeneralAlign
        | AlignLeft
        | AlignCenter
        | AlignRight
        | AlignFill
        | AlignJustify

    type VerticalAlignment =
        | AlignTop
        | AlignMiddle
        | AlignBottom

    /// Core covers wrap-text + the common alignment values. Indent level, text rotation,
    /// shrink-to-fit and reading order are a documented gap - see MAPPING.md.
    type AlignmentStyle =
        { Horizontal: HorizontalAlignment option
          Vertical: VerticalAlignment option
          WrapText: bool }

        static member Default =
            { Horizontal = None; Vertical = None; WrapText = false }

    /// A small set of named formats covering the vast majority of real spreadsheets, plus
    /// `Custom` as an escape hatch for any raw OOXML number format code. Named cases map to
    /// OOXML's built-in numFmtIds (which Excel recognizes and localizes for display);
    /// `Custom` codes are registered as custom numFmts (id >= 164).
    type NumberFormat =
        | General
        | Integer
        | TwoDecimal
        | Percentage
        | Currency
        | ShortDate
        /// Named `DateAndTime` rather than `DateTime` to avoid shadowing `System.DateTime`
        /// for anyone who has this module open alongside `System` - see MAPPING.md.
        | DateAndTime
        | Custom of formatCode: string

    /// Per-cell lock/hide flags. These only actually do anything once sheet-level
    /// protection (`SheetProtection`) is turned on for the worksheet - Excel ignores them
    /// otherwise. Every cell is `Locked = true` by default even without an explicit
    /// `CellStyle` at all, matching Excel's own default; `Hidden` (true = the cell's
    /// formula is hidden from the formula bar once protected) defaults to `false`.
    type CellProtection =
        { Locked: bool
          Hidden: bool }

        static member Default = { Locked = true; Hidden = false }

    type CellStyle =
        { Font: FontStyle option
          Fill: FillStyle option
          Border: BorderStyle option
          NumberFormat: NumberFormat option
          Alignment: AlignmentStyle option
          Protection: CellProtection option }

        static member Default =
            { Font = None
              Fill = None
              Border = None
              NumberFormat = None
              Alignment = None
              Protection = None }
