# SafeOpenXml

A typesafe F# DSL for building Excel workbooks, interpreted into calls against the
[DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) SDK. The DSL is a plain
data model (records/DUs with structural equality) — the interpreter (`Writer`) compiles it
to OOXML, and the reverse transform (`Reader`) parses an existing `.xlsx` back into the
same DSL.

See [MAPPING.md](MAPPING.md) for exactly which SpreadsheetML features map 1:1, which are
approximated, and which aren't modeled yet.

## Layout

- `src/SafeOpenXml` — the library.
  - `Reference.fs` — `CellRef` and `"A1"`-style address conversions.
  - `Styles.fs` — cell formatting: `Color`, `FontStyle`, `FillStyle`, `BorderStyle`,
    `AlignmentStyle`, `NumberFormat`, `CellProtection`, `CellStyle`.
  - `Validation.fs` — conditional formatting and data validation: `ComparisonOperator`
    (shared by both), `ConditionalFormatRule`, `ValidationKind`, `ValidationAlert`, and the
    `ConditionalFormatEntry`/`DataValidationEntry` records stored on `Worksheet`.
  - `Hyperlinks.fs` — `HyperlinkTarget` (external URL/`mailto:` vs. internal same-workbook
    reference) and the `HyperlinkEntry` record stored on `Worksheet`.
  - `Comments.fs` — `CommentEntry` (classic cell comments, i.e. current Excel's "Notes" -
    see MAPPING.md for the modern threaded-comments gap).
  - `Protection.fs` — `SheetProtection`, the sheet-level protection flags stored on
    `Worksheet` (pairs with `CellStyle.Protection` for per-cell locking).
  - `DefinedNames.fs` — `DefinedNameScope`/`DefinedNameEntry`, stored on `Workbook` rather
    than `Worksheet` - the one DSL concept that's genuinely workbook-level.
  - `PageSetup.fs` — print settings: `PageOrientation`, `PaperSize`, `PrintScaling`,
    `PageMargins`, and the `PageSetup` record stored on `Worksheet`.
  - `Tables.fs` — Excel Tables: `TableColumn`, `TableStyle`, and the `TableEntry` record
    stored as a list on `Worksheet` (a sheet can have several).
  - `Sparklines.fs` — in-cell mini-charts: `SparklineType`, `SparklineStyle`,
    `SparklineCell`, and the `SparklineGroupEntry` record stored as a list on `Worksheet`
    (a sheet can have several independently-styled groups).
  - `Model.fs` — `CellValue`, `Cell`, `Worksheet`, `Workbook`.
  - `Builders.fs` — ergonomic helpers: plain functional constructors (`cellA1`, ...) for
    the canonical model, plus the `SheetItem`/`CellEntry` types (each a single simple DU
    case with optional fields) and the `sheet` fold function - a small tree-shaped "AST
    for building a sheet" (rows of cells, plus sheet-level facts like column widths,
    merges, conditional formats, data validations, hyperlinks, comments, autofilter, and
    sheet protection) that mirrors how SpreadsheetML itself nests. `SheetDsl` is what you
    actually write against: `cell`/`row`/`autoFilter`/`conditionalFormat`/
    `dataValidation`/`hyperlink`/`comment` members with real optional parameters (`?col`,
    `?style`, `?index`, the data validation alert fields, `?tooltip`, `?author`) - no
    builder objects, no separate "styled" function, no `None`-noise for the common case.
    (`Protect` is the one `SheetItem` case with no smart constructor - `SheetProtection`
    is a plain record you build the usual F# way, `{ SheetProtection.Default with ... }`.)
  - `Interpreter/StyleRegistry.fs` — interns fonts/fills/borders/number formats into a
    shared OOXML stylesheet (internal).
  - `Interpreter/Writer.fs` — DSL → OOXML (internal).
  - `Interpreter/Reader.fs` — OOXML → DSL, the reverse transform (internal).
  - `Interpreter/CodeGen.fs` — DSL → F# *source text*: renders a `Workbook` back out as a
    self-contained `.fsx` script that rebuilds an equivalent file when run (internal).
  - `Api.fs` — the public `Workbook.save` / `saveToStream` / `load` / `loadFromStream` /
    `generateScript` entry points.
- `tests/SafeOpenXml.Tests` — one test per feature, each validating the produced file
  against the OOXML schema (`DocumentFormat.OpenXml.Validation.OpenXmlValidator`) and
  asserting an exact round trip back through the DSL. Each test also writes the workbook
  it builds to `Examples/<test name>/output.xlsx` (checked into the repo), so every
  feature has a real, openable `.xlsx` demonstrating it - a browsable gallery, not just
  assertions. Each scenario also gets an `Examples/<test name>/script.fsx` - see
  "Regenerating a file as F# source" below - which a separate, slower `Category=Slow` test
  group actually executes via `dotnet fsi` and verifies against the committed `.xlsx`.
- `samples/SafeOpenXml.Sample` — a small console app that builds a workbook, saves it,
  and reads it back.

## Quick start

```fsharp
open SafeOpenXml
open type SafeOpenXml.SheetDsl

let headerStyle =
    { CellStyle.Default with
        Font = Some { FontStyle.Default with Bold = true }
        Fill = Some { Color = Rgb(220uy, 220uy, 220uy) } }

let data =
    sheet
        "Sheet1"
        [ row [ cell (Text "Name", style = headerStyle)
                cell (Text "Amount", style = headerStyle) ]
          row [ cell (Text "Widgets")
                cell (Number 42.5, style = { CellStyle.Default with NumberFormat = Some TwoDecimal }) ]
          Freeze(1, 0) ]

workbook [ data ] |> Workbook.save "out.xlsx"

// Reverse transform:
let roundTripped = Workbook.load "out.xlsx"
```

`CellEntry` and `SheetItem`'s row case are each a single simple DU case with optional
fields (`Col`/`Index`) rather than separate "styled" or "explicit position" cases - `None`
means "the next column/row after the previous entry" (starting at 0), `Some n` jumps there
explicitly and sequential numbering resumes right after it. You don't construct the case
directly, though: `SheetDsl.cell`/`SheetDsl.row` are members with real optional
parameters (`?col`/`?style` on `cell`, `?index` on `row`) that hide the `None`s for
the common case - plain `let` functions can't have optional parameters in F#, which is why
this one bit of the DSL is a type. `open type SafeOpenXml.SheetDsl` (alongside `open SafeOpenXml`) brings `cell`/`row`
into scope unqualified, same as a module. Explicit column/row jumps go through the same
two members, just with the optional argument supplied: `cell (value, col = 2)` and
`row (cells, index = 4)`. `sheet` is the one fold that interprets the resulting item
list into the canonical `Worksheet` (the same relationship `Writer` has to OOXML). If you
already have cells pre-addressed by `CellRef` rather than grouped by row, `sheetOfCells`
builds a `Worksheet` directly from a flat `Cell list` instead.

Conditional formatting and data validation are `SheetItem`s too:

```fsharp
[ conditionalFormat (
    CellRef.ofA1 "A1",
    CellRef.ofA1 "A10",
    CellValueRule(GreaterThan, "100", None, { CellStyle.Default with Fill = Some { Color = Rgb(255uy, 199uy, 206uy) } })
  )
  dataValidation (CellRef.ofA1 "B1", CellRef.ofA1 "B10", ListValidation [ "Small"; "Medium"; "Large" ]) ]
```

See [MAPPING.md](MAPPING.md) for exactly which rule kinds of each are covered.

Defined names are workbook-level, so they attach to the `Workbook`, not a `Worksheet`:

```fsharp
workbook [ data ]
|> withDefinedNames
    [ definedName "TaxRate" "Sheet1!$A$1"
      sheetScopedDefinedName "Sheet1" "LocalTotal" "Sheet1!$A$2" ]
```

Print settings are a `SheetItem` too - `PageSetup` (the DU case) takes a plain
`PageSetup` record (the type), no smart constructor, same as `Protect`/`SheetProtection`.
`PrintArea` is a list of ranges (Excel supports several disjoint print rectangles per
sheet) - under the hood it's actually a hidden defined name, but `Writer`/`Reader`
translate transparently, so it reads and writes like any other `PageSetup` field:

```fsharp
[ PageSetup
    { PageSetup.Default with
        Orientation = Landscape
        Scaling = Some(FitToPage(1, 0)) // 1 page wide, unlimited tall
        PrintArea = [ (CellRef.ofA1 "A1", CellRef.ofA1 "D10") ]
        Header = Some "&C&\"Arial,Bold\"Quarterly Report"
        FirstHeader = Some "&CCover Page" // shown only on page 1
        EvenFooter = Some "&L&F" } ] // shown only on even pages
```

See [MAPPING.md](MAPPING.md) for what isn't modeled (totals-row/headerless tables, and a
handful of minor `pageSetup` attributes like print page order).

Tables are also a `SheetItem` - `Table` (the DU case) takes a plain `TableEntry` record
(the type), no smart constructor, same as `Protect`/`PageSetup`. Core doesn't synthesize
the header row's cell text for you, so it must already be there as ordinary cells - the
same way conditional formatting/autofilter/merges only describe metadata layered on top of
cells you've already placed:

```fsharp
sheet
    "Sheet1"
    [ row [ cell (Text "Item"); cell (Text "Quantity") ]
      row [ cell (Text "Widgets"); cell (Number 12.0) ]
      Table
          { TopLeft = CellRef.ofA1 "A1"
            BottomRight = CellRef.ofA1 "B2"
            Name = "Inventory"
            Columns = [ { Name = "Item"; CalculatedFormula = None }; { Name = "Quantity"; CalculatedFormula = None } ]
            Style = TableStyle.Default } ]
```

Structured references (`Table1[Column]`) need no special handling - they're just raw
formula text in a `Formula` cell, same as any other formula. See [MAPPING.md](MAPPING.md)
for what isn't modeled (totals row, headerless tables).

Sparklines follow the same shape - `SparklineGroup` (the DU case) takes a plain
`SparklineGroupEntry` record:

```fsharp
[ SparklineGroup
    { Style = { SparklineStyle.Default with Type = Column; ShowNegative = true }
      Sparklines =
        [ { Cell = CellRef.ofA1 "E1"; DataTopLeft = CellRef.ofA1 "A1"; DataBottomRight = CellRef.ofA1 "D1" } ] } ]
```

Sparklines are a Microsoft extension (living in the worksheet's `extLst`), not core
SpreadsheetML - unlike the rest of this library, schema validation alone can't confirm
real Excel renders one correctly, so treat this one with a bit more caution and verify in
real Excel before relying on it. See [MAPPING.md](MAPPING.md) for what isn't modeled
(axis settings, per-role colors beyond the main series color).

## Regenerating a file as F# source

Given a `Workbook` (typically one you just `Workbook.load`ed from an existing file),
`Workbook.generateScript` renders it back out as a self-contained `.fsx` script that
rebuilds an equivalent file when run - a code-generating counterpart to `Workbook.load`,
one level further than the reverse transform: instead of data, you get DSL *source text*.
It has no opinion on how the script locates the SafeOpenXml assembly, so you supply the
`#r` lines yourself:

```fsharp
let wb = Workbook.load "input.xlsx"

let referenceLines =
    [ "#r \"path/to/SafeOpenXml.dll\""
      "#r \"path/to/DocumentFormat.OpenXml.dll\"" ]

let script = Workbook.generateScript referenceLines "output.xlsx" wb
System.IO.File.WriteAllText("regenerate.fsx", script)
```

Running `dotnet fsi regenerate.fsx` produces `output.xlsx` - not byte-identical to the
original (zip metadata/timestamps differ) but structurally equivalent through the same
round-trip lens every other test in this repo uses. Generated code only ever mentions
fields that differ from `CellStyle.Default`/`BorderStyle.None`/etc., and only gives a
row/cell an explicit `index`/`col` where the source actually has a gap - see
`Interpreter/CodeGen.fs`. Every scenario under `tests/SafeOpenXml.Tests/Examples/` has a
committed `script.fsx` generated exactly this way; the `Category=Slow` test group is what
actually runs each one via `dotnet fsi` and checks it reproduces the committed `.xlsx`.

## Building and testing

```bash
dotnet build
dotnet test --filter "Category!=Slow"
dotnet run --project samples/SafeOpenXml.Sample
```

The default loop above skips the slow `Category=Slow` tests, which actually invoke
`dotnet fsi` on every generated `Examples/*/script.fsx` (multi-second process startup
each, so ~30-60s total) rather than just checking the generated source parses. Run those
explicitly, after the fast suite has populated the `.fsx` files at least once:

```bash
dotnet test --filter "Category=Slow"
```

Plain `dotnet test` (no filter) runs both groups.
