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
- AutoFilter: the range showing filter dropdown arrows (see the gap below on active
  filter criteria).
- Protection: per-cell `Locked`/`Hidden` flags (`CellStyle.Protection`, reused by
  conditional-formatting dxfs the same way the rest of `CellStyle` is) and sheet-level
  `SheetProtection` - a thin, direct pass-through of the OOXML `sheetProtection` flags
  (see that type's own doc comment for why: several of them are "true blocks the action",
  not "true allows it", and this deliberately never guesses a default for an unset flag).
  Passwords are hashed with the classic weak XOR algorithm on write for broad
  compatibility, and never round-trip back to plaintext (hashes aren't reversible) - see
  the gap below. Confirmed in real Excel: per-cell unlock behaves correctly (all other
  cells stay locked, as Excel itself defaults), and the password hash is correct - Excel
  accepts the intended password when unprotecting.
- Defined names: workbook-scoped or sheet-scoped named ranges/formulas/constants
  (`DefinedNameEntry`, stored on `Workbook` rather than `Worksheet` - the one DSL concept
  that's genuinely workbook-level, not per-sheet). `SheetScope` is expressed by sheet name
  and translated to/from OOXML's index-based `localSheetId` automatically.
- Print settings and page setup: orientation, paper size (`PaperSize`, a small named set
  plus `OtherPaperSize` for any other OOXML `ST_PaperSize` code), scaling (either a
  percentage or "fit to N pages wide by M tall", `0` meaning unconstrained in that
  dimension), margins, print area, and header/footer including its first-page/even-page
  variants (`PageSetup`, stored on `Worksheet`). Header/footer text is raw OOXML - Excel's
  own `&L`/`&C`/`&R`/`&P`/`&N`/`&D`/`&T`/`&F`/`&A` section/field codes embedded directly in
  one string per side, the same convention as OOXML's own `oddHeader`/`oddFooter`/etc.;
  setting `EvenHeader`/`EvenFooter` or `FirstHeader`/`FirstFooter` automatically sets the
  sibling `differentOddEven`/`differentFirst` flags that make Excel actually look at them.
  `FitToPage` scaling also sets the sibling `sheetPr/pageSetUpPr/@fitToPage` flag that
  tells Excel's print dialog which of `scale`/`fitToWidth`+`fitToHeight` to actually honor
  (both are always written regardless, so the file is self-describing either way - see
  `Interpreter/Writer.fs: pageSetupElement`). `PrintArea` (a list of ranges - Excel
  supports several disjoint print rectangles per sheet) is, under the hood, a reserved
  hidden sheet-scoped defined name (`_xlnm.Print_Area`) rather than a `pageSetup`
  attribute - `Writer`/`Reader` translate transparently, the same way `SheetScope`
  translates to/from OOXML's `localSheetId`, so callers never have to think about defined
  names for this.
- Tables: Excel Tables (`ListObject`s, the things structured references like
  `Table1[Column]` point at) over a range, with named columns (including an optional
  calculated-column formula) and a visual style reference (`TableEntry`/`TableColumn`/
  `TableStyle`, stored as a list on `Worksheet` - a sheet can have several). Structured
  references themselves need no special modeling: they're already just raw formula text,
  the same convention `CellValue.Formula`/`ConditionalFormatRule`/etc. all use. `Name`
  doubles as OOXML's separate `name` and `displayName` attributes (Core always writes the
  same value to both - see the gap below). See the gaps below for what a table can do in
  real Excel that Core doesn't model (totals row, headerless tables, active autofilter
  criteria on top of the table, custom table style *definitions*).
- Sparklines: groups of in-cell mini-charts (`SparklineGroupEntry`/`SparklineStyle`/
  `SparklineCell`, stored as a list on `Worksheet` - a sheet can have several
  independently-styled groups). Covers the commonly used subset: line/column/win-loss
  type, the main sparkline color, line weight, and the "highlight these points"
  high/low/first/last/negative toggles from Excel's Sparkline Design ribbon. Unlike every
  other feature above, this isn't core SpreadsheetML at all - it's a Microsoft extension
  living in the worksheet's `extLst` (`x14:sparklineGroups`, under a fixed, well-known
  extension URI real Excel files use for exactly this), so schema validation alone can't
  confirm real Excel actually renders it (the schema for `extLst`/`ext` content is
  deliberately open-ended) - see the gaps below, and note in the codebase pointing at
  where to manually verify in real Excel the same way Comments' VML rendering and
  SheetProtection's password hash were.
- Code generation: `Workbook.generateScript` renders any `Workbook` value (including one
  produced by `Workbook.load`) back out as an `.fsx` script that rebuilds a structurally
  equivalent file when run via `dotnet fsi` - covers every DSL construct above, since it's
  a direct pretty-printer over the same types (`Interpreter/CodeGen.fs`), not a separate
  model. Verified for every scenario in `tests/SafeOpenXml.Tests/Examples/`: each one's
  committed `script.fsx` is actually executed (not just generated) by the `Category=Slow`
  test group.

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
- **Workbook-level metadata.** Active sheet/tab selection and window sizing are not
  modeled (defined names are, as of `DefinedNameEntry` - see above).
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
- **AutoFilter criteria.** Only the filtered range itself is modeled - not any active
  filter conditions (`filterColumn` children: value lists, custom conditions, top-10,
  color/icon filters) a user or a foreign file may have configured on top of it. Reading
  a file with active criteria preserves the range but drops the criteria.
- **Password round-trip.** `SheetProtection.Password` always reads back as `None` - the
  hash isn't reversible, so re-saving a round-tripped protected file loses password
  enforcement unless the caller re-supplies the password. This is a deliberate consequence
  of not modeling raw hash/salt data as a DSL-level concept, not an oversight.
- **Newer, stronger password hash.** Only the classic weak hash is supported, not the
  modern salted-SHA-512 scheme (`algorithmName`/`hashValue`/`saltValue`/`spinCount`) newer
  Excel versions can also use - not modeled at all, on either the read or write side.
- **Print page order and other minor `pageSetup` attributes.** `pageOrder`
  (top-to-bottom vs. left-to-right), `firstPageNumber`, black-and-white/draft printing,
  and print resolution (`horizontalDpi`/`verticalDpi`) aren't modeled.
- **Table totals row.** Not modeled at all - `TableEntry` is always `headerRowCount="1"`,
  `totalsRowShown="0"`. Excel's totals-row dropdown (per-column sum/average/count/custom
  formula) and the extra worksheet row it occupies aren't represented; reading a table
  that has one drops the totals-row metadata (the underlying cells round-trip as ordinary
  `Cell`s, just outside what Core considers the table's own range).
- **Headerless tables.** Only `headerRowCount="1"` (a table with a header row, the
  overwhelmingly common case) is modeled - `headerRowCount="0"` isn't.
- **Table `name`/`displayName` as separate values.** OOXML allows a table's internal
  `name` and its formula-facing `displayName` to differ (rare in practice - Excel keeps
  them in sync unless edited some unusual way). `TableEntry.Name` is a single field
  written to both attributes on save, and read back from `name` - a foreign file where
  they genuinely differ loses the `displayName`.
- **Table style *definitions*.** `TableStyle.Name` is a reference to a style by name
  (a built-in like `"TableStyleMedium2"`, or a custom one defined elsewhere in the
  workbook) - Core doesn't model custom table style *definitions* themselves (the
  `tableStyles`/`dxf`-based colors and fonts a workbook can define), only the reference.
- **Table-specific autofilter criteria.** Every table always gets an `autoFilter` element
  matching its own range (for the header row's dropdown arrows), but - same gap as the
  standalone `AutoFilter` feature - only the arrows are modeled, not any active filter
  conditions on top of them.
- **Table/column name uniqueness across the workbook.** `Writer` validates that a single
  table's column count matches its range width and that its own column names are unique
  (both would otherwise produce a file Excel refuses to open cleanly) - it does not
  validate that table names or ids are unique *across* the whole workbook, the same way
  sheet names and defined names aren't cross-checked either.
- **Sparkline axis settings.** Manual min/max, date-axis mode, and whether the horizontal
  axis line itself is shown (`minAxisType`/`maxAxisType`/`manualMin`/`manualMax`/
  `dateAxis`/`displayXAxis`) aren't modeled - a group using them keeps its data/style on
  round trip but loses the axis configuration.
- **Sparkline per-role colors.** Only the main series color (`SparklineStyle.Color`) is
  modeled. Excel's separate colors for negative points, the axis line, and each marker
  role (regular/first/last/high/low) aren't - they always come through as Excel's own
  automatic choices, even if a foreign file set them explicitly.
- **Sparkline empty-cell/hidden-cell handling and right-to-left.** `displayEmptyCellsAs`
  (gap/zero/connect), `displayHidden` (plot hidden rows/columns), and `rightToLeft` aren't
  modeled.
- **Sparkline data range sheet.** `SparklineCell`'s data range is always assumed to be on
  the same sheet as the sparkline itself - a foreign file with a sparkline pointing at
  another sheet's data has that sheet qualifier silently discarded on read (the range
  itself still parses, just interpreted against the sparkline's own sheet).

## Out of scope for Core (candidates for a future extension)

These are real SpreadsheetML features with no DSL representation at all, by design — Core
was scoped to the cell/style/layout fundamentals first (see the assistant's initial scoping
proposal). Each would be its own reasonably-sized module to add later, verified against
real files the same way Core was:

- Charts
- Images / drawings
- Pivot tables
- Workbook-level protection (protecting the workbook structure itself - sheet ordering,
  visibility - as distinct from the per-sheet protection Core already models)
- Macros / VBA

## A note on style interning

The DSL's `CellStyle` is a plain F# record, so two cells that specify *structurally equal*
styles get free deduplication into a single shared stylesheet entry when writing (mirroring
what Excel itself does) — see `Interpreter/StyleRegistry.fs`. There is no equivalent
concept of "named cell styles" (the `cellStyleXfs`/`cellStyles` collections, e.g. Excel's
"Good/Bad/Neutral" named styles) — Core only ever emits/reads a single implicit "Normal"
named style, and per-cell formatting goes through direct cell formats (`cellXfs`).
