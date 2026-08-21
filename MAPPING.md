# DSL ↔ OOXML mapping

SafeOpenXml's DSL (`src/SafeOpenXml/{Reference,Styles,Model}.fs`) aims to map 1:1 onto
SpreadsheetML wherever the DSL models a feature at all. This document lists every place
where that mapping is inexact, lossy, or simply not implemented yet in Core, so you know
what to expect from a round trip (`Workbook.save` then `Workbook.load`) and what would
need to be added to close the gap.

## Modeled faithfully (1:1 or as close as makes sense)

- Cell addressing (`CellRef` ↔ `"A1"` references), including the bijective base-26 column
  scheme.
- Cell values: text (as shared strings), numbers, booleans, dates (OLE Automation serial
  dates via `DateTime.ToOADate`/`FromOADate`), and formulas (raw formula text plus an
  optional cached numeric result).
- Font: name, size, bold, italic, underline, strikethrough, color.
- Fill: solid color only (see gaps below).
- Border: left/right/top/bottom, a core set of named line styles plus a `BorderLineStyle.Other`
  escape hatch that preserves any OOXML border style by name.
- Alignment: horizontal, vertical, wrap text.
- Number formats: `General`, `Integer`, `TwoDecimal`, `Percentage`, `Currency`, `ShortDate`,
  `DateAndTime`, plus `Custom` for any raw format code.
- Column widths, row heights, merged cell ranges, frozen panes (leading rows/columns).
- Colors: `Rgb` fully round-trips. `Indexed` and `Theme` (with optional tint) are preserved
  losslessly as opaque references for read-then-write round trips, but Core does not
  *resolve* them to actual RGB values (that requires parsing the workbook's theme part,
  which Core does not model) — see below.
- Conditional formatting: `CellValueRule` (comparison operators, incl. `Between`/
  `NotBetween`), `FormulaRule`, 2-color and 3-color scales, a single-color data bar
  (automatic min/max thresholds), and duplicate/unique-value highlighting. Styles for
  these reuse `CellStyle` directly, written into the stylesheet's `dxfs` (differential
  formats) collection — see `Interpreter/StyleRegistry.fs: InternDxf`.
- Data validation: whole-number/decimal/text-length rules (with the same comparison
  operators as conditional formatting), inline dropdown lists, range-sourced dropdown
  lists, and arbitrary custom-formula rules — plus the optional input prompt and error
  alert messages.
- Hyperlinks: external (any URL, including `mailto:`) and internal (same-workbook,
  `"Sheet2!A1"`-style or a defined name) targets, over a single cell or a range, with an
  optional tooltip and an optional fallback display label.
- Comments: classic cell comments (what current Excel's UI calls "Notes" - see the gap
  below on modern threaded comments) with author and text, deduplicating authors the same
  way shared strings are deduplicated. Written with the accompanying legacy VML drawing
  part real Excel-authored files pair every classic comment with, for the on-cell
  indicator and hover-box position — confirmed rendering correctly (indicator, hover
  popup, position) in real Excel.

## Known gaps (documented, not silently "supported")

- **Rich text runs.** A `Text` cell is always one uniformly-styled string. OOXML supports
  multiple runs with different fonts/colors inside a single cell; reading such a cell
  concatenates all runs' text (via `InnerText`) and discards the per-run formatting.
- **Fill patterns.** Only `patternType="solid"` is modeled. Pattern fills (stripes,
  checkerboards, gradients) are not represented — reading a cell with a non-solid pattern
  yields `Fill = None` for that cell.
- **Diagonal borders.** Only the four edge borders are modeled; diagonal cell borders are
  not.
- **Uncommon border line styles.** `Thin`/`Medium`/`Thick`/`Dashed`/`Dotted`/`Double`/`Hair`
  are named cases. The remaining OOXML styles (`mediumDashed`, `dashDot`, `mediumDashDot`,
  `dashDotDot`, `mediumDashDotDot`, `slantDashDot`) come through as `Other "<rawName>"` —
  round-trippable, but not first-class.
- **Theme/indexed colors are not resolved.** `Theme(themeIndex, tint)` and
  `Indexed(paletteIndex)` preserve the *reference* faithfully, but Core cannot tell you
  what RGB color that actually renders as (needs the theme XML part, which isn't parsed).
  You can construct a `Theme`/`Indexed` color yourself only if you already know the target
  workbook's palette.
- **Alignment extras.** Indent level, text rotation, shrink-to-fit, justify-last-line, and
  reading order are not modeled.
- **Number format built-ins beyond the named set.** On read, numFmtIds `3`, `4`, `9`, `37`–`40`,
  and `49` are recognized and preserved as `Custom "<code>"`; any *other* unrecognized
  built-in numFmtId currently degrades to `None` (General) on read. Easy to extend — see
  `Interpreter/Reader.fs: otherBuiltinFormatCodes`.
- **Formula fidelity.** Formula text itself round-trips exactly. Shared-formula grouping
  and array-formula (`t="array"`) metadata are not preserved as such — Core always writes
  a plain, ungrouped formula per cell.
- **Cell error values** (`#DIV/0!`, `#N/A`, ...) round-trip as `Text` of the literal error
  string rather than a dedicated error case.
- **Locale-dependent built-in formats.** `Currency` writes the fixed format code
  `"$"#,##0.00`. `ShortDate`/`DateAndTime` use OOXML's built-in numFmtIds 14/22, whose
  on-screen rendering is locale-dependent in Excel itself — this mirrors Excel's own
  behavior rather than being a SafeOpenXml limitation.
- **Hidden/outlined rows and columns.** Not modeled — hidden state and outline level are
  dropped.
- **Split panes.** Only the common "frozen leading rows/columns" case is modeled; the
  independent-scroll "Split" pane state is not.
- **Workbook-level metadata.** Defined names/named ranges, active sheet/tab selection, and
  window sizing are not modeled.
- **Conditional formatting rule kinds.** Icon sets, "top/bottom N (or %)", above-average,
  time-period, and the text/blank/error-contains rule kinds are not modeled — a `cfRule`
  of one of these kinds is silently dropped on read (per this project's "drop what isn't
  modeled" round-trip philosophy) rather than causing a failure.
- **Conditional formatting ranges.** Each `ConditionalFormatEntry` covers exactly one
  rectangular range. OOXML's `sqref` actually accepts a space-separated list of ranges per
  rule; Core always writes/reads just the first one.
- **Data bar/color scale details.** Excel's richer data-bar options (gradient vs. solid
  fill, explicit min/max thresholds instead of automatic, axis position, border, negative-
  value color, direction) aren't modeled — `DataBarRule` only covers a single-color bar
  with automatic thresholds, matching Excel's own "quick" default.
- **Data validation `Date`/`Time` types.** Not modeled; a `dataValidation` of type `date` or
  `time` round-trips as `CustomValidation` of its raw formula text instead of a dedicated
  case.
- **Data validation named-range list sources.** A dropdown sourced from a defined name
  (rather than a literal range or an inline list) degrades to a single-item literal list
  on read.
- **Comment position/size.** Only the comment's cell, author, and text are modeled - the
  VML anchor's box size and on-screen position are fixed defaults (matching a freshly
  inserted Excel comment), not configurable, and not read back from an existing file's
  VML (which isn't parsed on read at all - see `Interpreter/Reader.fs: readComments`).
- **Threaded comments.** Modern Excel's actual "Comments" (2016+, `@mentions`, replies,
  resolved state) are a completely different part format (`WorksheetThreadedCommentsPart`)
  not modeled here at all - what this library calls a `CommentEntry` is what current
  Excel's UI now labels a "Note".

## Out of scope for Core (candidates for a future extension)

These are real SpreadsheetML features with no DSL representation at all, by design — Core
was scoped to the cell/style/layout fundamentals first (see the assistant's initial scoping
proposal). Each would be its own reasonably-sized module to add later, verified against
real files the same way Core was:

- Charts
- Images / drawings
- Pivot tables
- Sheet/workbook protection
- Print settings and page setup
- Tables (`ListObject`s / structured references) and autofilter
- Sparklines
- Macros / VBA

## A note on style interning

The DSL's `CellStyle` is a plain F# record, so two cells that specify *structurally equal*
styles get free deduplication into a single shared stylesheet entry when writing (mirroring
what Excel itself does) — see `Interpreter/StyleRegistry.fs`. There is no equivalent
concept of "named cell styles" (the `cellStyleXfs`/`cellStyles` collections, e.g. Excel's
"Good/Bad/Neutral" named styles) — Core only ever emits/reads a single implicit "Normal"
named style, and per-cell formatting goes through direct cell formats (`cellXfs`).
