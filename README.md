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
  - `Api.fs` — the public `Workbook.save` / `saveToStream` / `load` / `loadFromStream`
    entry points.
- `tests/SafeOpenXml.Tests` — one test per feature, each validating the produced file
  against the OOXML schema (`DocumentFormat.OpenXml.Validation.OpenXmlValidator`) and
  asserting an exact round trip back through the DSL. Each test also writes the workbook
  it builds to `Examples/<test name>/output.xlsx` (checked into the repo), so every
  feature has a real, openable `.xlsx` demonstrating it - a browsable gallery, not just
  assertions.
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

## Building and testing

```bash
dotnet build
dotnet test
dotnet run --project samples/SafeOpenXml.Sample
```
