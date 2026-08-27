# Kookerella.FsOpenXmlDsl

A typesafe F# DSL for building Excel workbooks, interpreted into calls against the
[DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) SDK. The DSL is a plain
data model (records/DUs with structural equality) — the interpreter (`Writer`) compiles it
to OOXML, and the reverse transform (`Reader`) parses an existing `.xlsx` back into the
same DSL.

See [MAPPING.md](MAPPING.md) for exactly which SpreadsheetML features map 1:1, which are
approximated, and which aren't modeled yet.

**This round-trips in both directions**, which most Excel libraries (EPPlus, ClosedXML,
NPOI, ...) don't: they give you an imperative API to build a workbook from scratch or mutate
an existing one, but no way to turn an *existing* file back into readable source. Here,
`Reader` parses a real `.xlsx`/`.xlsm` back into the same DSL, and `Workbook.generateScript`
(F#) / `CsCodeGen.Generate` (C#) go one step further and render that model back out as a
self-contained script that rebuilds an equivalent file - a decompiler for spreadsheets, not
just a writer. Two more surfaces, `Xml.ofWorkbook`/`Xml.toWorkbook` (see ["## XML"](#xml)
below) and `Json.ofWorkbook`/`Json.toWorkbook` (see ["## JSON"](#json) below), do the same
translation to/from plain XML or JSON against a real schema - for a caller who'd rather
generate or consume data than write code at all, e.g. an XSLT pipeline producing a report.
`Kookerella.FsOpenXmlDsl.Mcp` exposes all four directions as MCP tools
(`generate_fsharp_script`/`generate_csharp_script`/`generate_xml`/`create_workbook_from_xml`/
`generate_json`/`create_workbook_from_json`) for an AI agent, and as `fsopenxmldsl-mcp
convert`/`build` CLI commands for anyone else - try it on any spreadsheet you already have,
no code required:

```bash
dotnet tool install -g Kookerella.FsOpenXmlDsl.Mcp
fsopenxmldsl-mcp convert your-file.xlsx --lang csharp
```

## Layout

- `src/Kookerella.FsOpenXmlDsl` — the library.
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
    `Worksheet` (pairs with `CellStyle.Protection` for per-cell locking), and
    `WorkbookProtection`, the workbook-level structure/window protection flags stored on
    `Workbook`.
  - `DefinedNames.fs` — `DefinedNameScope`/`DefinedNameEntry`, stored on `Workbook` rather
    than `Worksheet` - the one DSL concept that's genuinely workbook-level.
  - `PageSetup.fs` — print settings: `PageOrientation`, `PaperSize`, `PrintScaling`,
    `PageMargins`, and the `PageSetup` record stored on `Worksheet`.
  - `Tables.fs` — Excel Tables: `TableColumn`, `TableStyle`, and the `TableEntry` record
    stored as a list on `Worksheet` (a sheet can have several).
  - `Sparklines.fs` — in-cell mini-charts: `SparklineType`, `SparklineStyle`,
    `SparklineCell`, and the `SparklineGroupEntry` record stored as a list on `Worksheet`
    (a sheet can have several independently-styled groups).
  - `Charts.fs` — column/bar/line/pie charts: `ChartType`, `ChartSeries`, and the
    `ChartEntry` record stored as a list on `Worksheet` (a sheet can have several).
  - `Images.fs` — raster images: `ImageFormat` and the `ImageEntry` record (raw file
    bytes plus a cell-range anchor) stored as a list on `Worksheet`.
  - `PivotTables.fs` — `PivotAggregation` and the `PivotTableEntry` record (source range,
    row/column/value fields, an anchor cell) stored as a list on `Worksheet`.
  - `Model.fs` — `CellValue`, `Cell`, `Worksheet`, `Workbook` (including `Workbook.
    VbaProject`, a macro-enabled workbook's raw `vbaProject.bin` bytes - see its own doc
    comment; there's no dedicated `Macros.fs` since it's a single opaque field, not a new
    type).
  - `Xml.fs` / `Xml.xsd` — the XML surface: `Xml.toWorkbook`/`Xml.ofWorkbook` translate a
    `Workbook` to/from an `XElement` tree, and `Xml.schemaSet()` loads the paired schema
    (embedded in the assembly as a resource) for validating either direction. See
    ["## XML"](#xml) below.
  - `Json.fs` — the JSON surface: `Json.toWorkbook`/`Json.ofWorkbook` translate a `Workbook`
    to/from a `System.Text.Json.Nodes.JsonObject` tree, covering the same
    worksheet/workbook-level feature set `Xml.fs` does. Schema validation
    (`Json.schema.json`) is test-suite only, not a public API - see ["## JSON"](#json)
    below.
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
  - `Interpreter/ChartWriter.fs` / `ChartReader.fs` — charts' own DSL ↔ DrawingML/ChartML
    translation, split out from `Writer.fs`/`Reader.fs` given how much larger that one
    feature's OOXML surface is than everything else combined (internal).
  - `Interpreter/ImageWriter.fs` / `ImageReader.fs` — images' own DSL ↔ DrawingML
    translation (internal).
  - `Interpreter/DrawingWriter.fs` / `DrawingReader.fs` — own the one `DrawingsPart`/
    `<drawing>` relationship a worksheet gets when it has charts and/or images, since both
    features share that one drawing canvas rather than each managing their own (internal).
  - `Interpreter/PivotTableWriter.fs` / `PivotTableReader.fs` — pivot tables' own group-by
    + aggregate engine plus DSL ↔ OOXML translation (`pivotCacheDefinition`/
    `pivotCacheRecords`/`pivotTableDefinition`), split out from `Writer.fs`/`Reader.fs` the
    same way charts and images are (internal).
  - `Interpreter/Writer.fs` — DSL → OOXML (internal).
  - `Interpreter/Reader.fs` — OOXML → DSL, the reverse transform (internal).
  - `Interpreter/CodeGen.fs` — DSL → F# *source text*: renders a `Workbook` back out as a
    self-contained `.fsx` script that rebuilds an equivalent file when run (internal).
  - `Api.fs` — the public `Workbook.save` / `saveToStream` / `load` / `loadFromStream` /
    `generateScript` entry points.
- `tests/Kookerella.FsOpenXmlDsl.Tests` — one test per feature, each validating the produced file
  against the OOXML schema (`DocumentFormat.OpenXml.Validation.OpenXmlValidator`) and
  asserting an exact round trip back through the DSL. Each test also writes the workbook
  it builds to `Examples/<test name>/output.xlsx` (checked into the repo), so every
  feature has a real, openable `.xlsx` demonstrating it - a browsable gallery, not just
  assertions. Each scenario also gets an `Examples/<test name>/script.fsx` - see
  "Regenerating a file as F# source" below - which a separate, slower `Category=Slow` test
  group actually executes via `dotnet fsi` and verifies against the committed `.xlsx`, and
  an `Examples/<test name>/workbook.xml` - the same workbook through `Xml.ofWorkbook`,
  validated against `Xml.xsd` at generation time (see "## XML" below) - and an
  `Examples/<test name>/workbook.json` - the same workbook through `Json.ofWorkbook`,
  validated against `Json.schema.json` at generation time (see "## JSON" below) - so one
  folder always has four views of the same example: the real file, the F# source that
  rebuilds it, and the XML/JSON that also rebuild it.
  `Assets/` holds the one test fixture too large to inline as a base64 literal like every
  other binary fixture in `Tests.fs` - a real `vbaProject.bin` extracted from a workbook
  actually saved by Excel, used by the macro example.
- `samples/Kookerella.FsOpenXmlDsl.Sample` — a small console app that builds a workbook, saves it,
  and reads it back.
- `src/Kookerella.CsOpenXmlDsl` — an idiomatic, immutable, fluent C# wrapper over this
  library, for callers who'd rather not touch F# discriminated unions/option types
  directly. Now covers every feature this library models at the worksheet/workbook level -
  see its own README for scope and an example. `tests/Kookerella.CsOpenXmlDsl.Tests` is its
  own C# xUnit suite, exercising the wrapper the way a real C# caller would rather than
  reusing the F# test project.
- `src/Kookerella.FsOpenXmlDsl.Mcp` — a local MCP (Model Context Protocol) server exposing
  this library's read/write/code-generation/XML/JSON capabilities as tools any
  MCP-compatible AI agent can call directly, and the same conversion capability as plain
  `fsopenxmldsl-mcp convert`/`build` CLI commands for anyone not going through an MCP
  client - see its own README for the tool list and how to configure it.

## Quick start

```fsharp
open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

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
this one bit of the DSL is a type. `open type Kookerella.FsOpenXmlDsl.SheetDsl` (alongside `open Kookerella.FsOpenXmlDsl`) brings `cell`/`row`
into scope unqualified, same as a module. Explicit column/row jumps go through the same
two members, just with the optional argument supplied: `cell (value, col = 2)` and
`row (cells, index = 4)`. `sheet` is the one fold that interprets the resulting item
list into the canonical `Worksheet` (the same relationship `Writer` has to OOXML). If you
already have cells pre-addressed by `CellRef` rather than grouped by row, `sheetOfCells`
builds a `Worksheet` directly from a flat `Cell list` instead.

**A `Formula` cell is `Formula(expression, cachedValue: float option)` - this library never
evaluates formulas itself, so `cachedValue` is the only number that will ever exist for that
cell until something else computes one.** Real Excel recalculates on open and overwrites it,
so leaving it `None` is fine if a human always opens the result in Excel first. It's *not*
safe for a headless pipeline - e.g. generating a workbook and piping it straight into a PDF
converter, another automated reader, or anything else that never opens it in real Excel.
Whether that downstream step shows a correct number, a blank, or a stale one depends
entirely on whether *it* happens to have its own formula engine; some do (Aspose.Cells,
Syncfusion, GemBox, real Excel via COM), many lighter-weight or headless converters don't and
will just render whatever's already in the cell. Since you already have the numbers that fed
into the formula, always pass the real result as `cachedValue` for anything that isn't
guaranteed to pass through Excel first - it costs nothing and sidesteps the problem
entirely, since a downstream reader with no evaluator at all can still show a correct value
someone else already computed.

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

Workbook-level protection (as distinct from a `Worksheet`'s own `SheetProtection`) is
also workbook-level, same pipe-friendly shape:

```fsharp
workbook [ data ]
|> withProtection { WorkbookProtection.Default with LockStructure = Some true }
```

`withDefinedNames`/`withProtection` compose - pipe both onto the same `workbook [...]`.

Macros are also workbook-level, same pipe-friendly shape - `withVbaProject` takes the raw
bytes of an existing `vbaProject.bin` (extracted from an `.xlsm` you already have, e.g. via
`System.IO.Compression.ZipFile`, or authored in Excel's VBA editor and harvested the same
way). Core doesn't decode, generate, or otherwise understand VBA source - it embeds and
reads back exactly the bytes you give it, the same "opaque payload" treatment
`ImageEntry.Data` gets for raster images:

```fsharp
workbook [ data ]
|> withVbaProject (System.IO.File.ReadAllBytes("vbaProject.bin"))
```

Save the result with an `.xlsm` path - `Workbook.save`/`saveToStream` automatically switch
the file's own declared content type to Excel's macro-enabled kind whenever a `VbaProject`
is present, but real Excel also expects the `.xlsm` extension to trust and run macros at
all. See [MAPPING.md](MAPPING.md) for what isn't modeled (authoring macro source, and the
one case where the default sheet/workbook codenames Core writes won't match what a macro's
original author intended).

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

Charts are the same shape too - `EmbeddedChart` (not bare `Chart`, which collides with
the OOXML SDK's own type - see `Builders.fs`) takes a plain `ChartEntry` record. A
series' `Name` is a reference to the cell that names it (its column header, typically),
live-updating the same way a real Excel chart's series name does - not a static copy:

```fsharp
[ EmbeddedChart
    { Type = ChartColumn
      Title = Some "Sales by Quarter"
      CategoriesTopLeft = CellRef.ofA1 "A2"
      CategoriesBottomRight = CellRef.ofA1 "A4"
      Series = [ { Name = CellRef.ofA1 "B1"; ValuesTopLeft = CellRef.ofA1 "B2"; ValuesBottomRight = CellRef.ofA1 "B4" } ]
      ShowLegend = true
      TopLeftAnchor = CellRef.ofA1 "E1"
      BottomRightAnchor = CellRef.ofA1 "L15" } ]
```

Unlike Sparklines, charts are core, fully schema-driven DrawingML/ChartML - built from
typed OOXML SDK classes the same way every other feature is, not an extension mechanism.
See [MAPPING.md](MAPPING.md) for what isn't modeled (chart kinds beyond column/bar/line/
pie, per-series styling, stacked grouping).

Images are anchored the same way - `EmbeddedImage` takes a plain `ImageEntry` record.
`Data` is just the image file's own raw bytes (read it with `System.IO.File.ReadAllBytes`,
for example) - this DSL doesn't decode or re-encode anything, only embeds and hands back
exactly what you give it:

```fsharp
[ EmbeddedImage
    { Data = System.IO.File.ReadAllBytes("logo.png")
      Format = Png
      TopLeftAnchor = CellRef.ofA1 "A1"
      BottomRightAnchor = CellRef.ofA1 "C6" } ]
```

A worksheet's charts and images share one drawing canvas under the hood (Excel only gives
a sheet one at all), which is transparent to you as a caller - just add both kinds of
`SheetItem` to the same sheet. See [MAPPING.md](MAPPING.md) for what isn't modeled (formats
beyond PNG/JPEG/GIF/BMP, free-floating position, cropping, linked-not-embedded images).

Pivot tables are also a `SheetItem` - `EmbeddedPivotTable` (not bare `PivotTable`, again
for naming consistency with `EmbeddedChart`/`EmbeddedImage`) takes a plain `PivotTableEntry`
record. Unlike every other feature, this one does real work at write time rather than a
pure translation: it groups the source range by `RowField` (and `ColumnField`, if given),
aggregates `ValueField`, and writes both a real Excel pivot cache and the resulting grid of
computed cells:

```fsharp
[ EmbeddedPivotTable
    { SourceSheet = None // defaults to this sheet; can name another
      SourceTopLeft = CellRef.ofA1 "A1"
      SourceBottomRight = CellRef.ofA1 "C5"
      RowField = "Region"
      ColumnField = Some "Quarter"
      ValueField = "Sales"
      Aggregation = PivotSum
      ValueCaption = Some "Total Sales"
      TopLeftAnchor = CellRef.ofA1 "E1" } ]
```

The source range's first row must be plain `Text` header cells naming each field. This is
deliberately scoped to what a single field per axis can express - one row field, at most
one column field, one value field, Tabular layout, grand totals only - see
[MAPPING.md](MAPPING.md) for the reasoning and what a richer pivot table (nested fields,
multiple value fields, page filters) would need instead.

## Regenerating a file as F# source

Given a `Workbook` (typically one you just `Workbook.load`ed from an existing file),
`Workbook.generateScript` renders it back out as a self-contained `.fsx` script that
rebuilds an equivalent file when run - a code-generating counterpart to `Workbook.load`,
one level further than the reverse transform: instead of data, you get DSL *source text*.
It has no opinion on how the script locates the FsOpenXmlDsl assembly, so you supply the
`#r` lines yourself:

```fsharp
let wb = Workbook.load "input.xlsx"

let referenceLines =
    [ "#r \"path/to/Kookerella.FsOpenXmlDsl.dll\""
      "#r \"path/to/DocumentFormat.OpenXml.dll\"" ]

let script = Workbook.generateScript referenceLines "output.xlsx" wb
System.IO.File.WriteAllText("regenerate.fsx", script)
```

Running `dotnet fsi regenerate.fsx` produces `output.xlsx` - not byte-identical to the
original (zip metadata/timestamps differ) but structurally equivalent through the same
round-trip lens every other test in this repo uses. Generated code only ever mentions
fields that differ from `CellStyle.Default`/`BorderStyle.None`/etc., and only gives a
row/cell an explicit `index`/`col` where the source actually has a gap - see
`Interpreter/CodeGen.fs`. Every scenario under `tests/Kookerella.FsOpenXmlDsl.Tests/Examples/` has a
committed `script.fsx` generated exactly this way; the `Category=Slow` test group is what
actually runs each one via `dotnet fsi` and checks it reproduces the committed `.xlsx`.

## XML

`Xml.toWorkbook`/`Xml.ofWorkbook` (in `Xml.fs`) are a third way in and out of the DSL,
alongside writing F#/C# directly and code generation: plain XML, against a real schema
(`Xml.xsd`, embedded in the assembly). This exists for a caller who'd rather generate or
consume data than write code at all. Two concrete uses:

- **Build an `.xlsx` from XML a transform engine already produces** - an XSLT pipeline (or
  any templating that emits XML) can target Excel directly, without learning the OOXML
  schema or this library's own API.
- **Convert an existing `.xlsx` to XML for version control** - `.xlsx` is a binary ZIP, so
  `git diff` on one is useless; converting to XML first makes a real, human-readable diff
  possible. `Xml.ofWorkbook`'s output is deterministically ordered (sorted by cell position,
  or by name for defined names) regardless of the order the underlying `Workbook`'s lists
  happen to be in, so a genuine content change produces a small, isolated diff rather than a
  spurious one from rows/rules getting reshuffled between runs.

```fsharp
open System.Xml.Linq

// XML -> Workbook -> .xlsx
let wb = XElement.Load "report.xml" |> Xml.toWorkbook
Workbook.save "report.xlsx" wb

// .xlsx -> Workbook -> XML
let xml = Workbook.load "report.xlsx" |> Xml.ofWorkbook
xml.Save "report.xml"
```

A discriminated union case becomes an XML element named after the case (camelCased) when
it carries data of its own, or an attribute *value* (also camelCased) when it's one of
several parameterless alternatives - e.g. a cell's value:

```xml
<cell ref="B2">
  <number>42.5</number>
  <style>
    <numberFormat kind="currency" />
  </style>
</cell>
```

A richer example - `ValidationKind`'s six cases follow the same convention, and
`ValidationAlert`'s fields are written as attributes directly on `<dataValidation>` itself
rather than nested:

```xml
<dataValidation topLeft="A2" bottomRight="A2" errorTitle="Invalid quantity"
                errorMessage="Quantity must be a positive whole number.">
  <wholeNumberValidation operator="greaterThan" formula1="0" />
</dataValidation>
```

`ConditionalFormatRule`'s seven cases follow the same convention too, nesting a full
`CellStyle` where the rule needs one - note `<fill>` holds `<rgb>`/`<indexed>`/`<theme>`
directly, with no extra wrapper element:

```xml
<conditionalFormat topLeft="A1" bottomRight="A3">
  <cellValueRule operator="greaterThan" formula1="100">
    <style>
      <fill>
        <rgb r="255" g="199" b="206" />
      </fill>
    </style>
  </cellValueRule>
</conditionalFormat>
```

A `Chart`'s `Series` list needs its own wrapper element (`<series>`) distinct from each
item's own element name (`<s>`), to avoid a real ambiguity XML has and JSON doesn't - a
list has no shape of its own in XML the way a JSON array does, so the container and its
items need different names or a reader can't tell where the list starts:

```xml
<chart type="column" title="Sales by Quarter" showLegend="true"
       anchorTopLeft="E1" anchorBottomRight="L15">
  <categories topLeft="A2" bottomRight="A4" />
  <series>
    <s name="B1" valuesTopLeft="B2" valuesBottomRight="B4" />
    <s name="C1" valuesTopLeft="C2" valuesBottomRight="C4" />
  </series>
</chart>
```

An Excel `Table` shows the more usual case for that same wrapper/item split - `columns`
already has a natural singular (`column`), so no `<s>`-style workaround is needed:

```xml
<table topLeft="A1" bottomRight="B4" name="Calc">
  <columns>
    <column name="Qty" />
    <column name="Doubled" calculatedFormula="Calc[Qty]*2" />
  </columns>
  <style name="TableStyleLight9" showFirstColumn="true" showLastColumn="true"
         showColumnStripes="true" />
</table>
```

A `SparklineGroup`'s `Color` field wraps in its own `<color>` child element, same
convention `CellStyle`'s font/fill use:

```xml
<sparklineGroup>
  <style type="column" lineWeight="1.5" showNegative="true">
    <color>
      <rgb r="0" g="112" b="192" />
    </color>
  </style>
  <sparklines>
    <sparkline cell="E1" dataTopLeft="A1" dataBottomRight="D1" />
  </sparklines>
</sparklineGroup>
```

A `PivotTable` is the flattest shape here - just attributes, no nested elements at all.
Note this only carries the *description* through: loading one via `Xml.toWorkbook` doesn't
re-run the aggregation, unlike everything else this schema covers:

```xml
<pivotTable sourceSheet="Data" sourceTopLeft="A1" sourceBottomRight="C9"
            rowField="Region" columnField="Quarter" valueField="Sales"
            aggregation="average" valueCaption="Avg Sales" anchorTopLeft="F1" />
```

An `Image`'s raw bytes are the element's own base64 text content, the same convention
`vbaProject` below uses:

```xml
<image format="gif" topLeft="A1" bottomRight="D6">R0lGODlhAQABAIAAAAAAAP...</image>
```

A `Hyperlink`'s `Target` nests the same way `ValidationKind`/`ConditionalFormatRule` do:

```xml
<hyperlink topLeft="A1" bottomRight="A1" tooltip="Visit site">
  <externalHyperlink>https://example.com</externalHyperlink>
</hyperlink>
<hyperlink topLeft="A2" bottomRight="B3" display="Go to top">
  <internalHyperlink>Sheet1!A1</internalHyperlink>
</hyperlink>
```

A `Comment`'s text is also the element's own content, not an attribute - `author` is
simply omitted when empty rather than written as `author=""`:

```xml
<comment cell="A1" author="Alex">Check this figure</comment>
<comment cell="A2">Unnamed author</comment>
```

Sheet and workbook protection are both flat attribute bags - no nested elements needed,
since none of `SheetProtection`/`WorkbookProtection`'s fields are structured data:

```xml
<protection password="hunter2" sheet="true" formatCells="true" sort="true" autoFilter="true" />
```

```xml
<workbook>
  <sheets>...</sheets>
  <protection password="hunter2" lockStructure="true" />
</workbook>
```

`PageSetup` shows the mixed-DU convention again - `PaperSize`'s named cases become a
`kind` attribute, the same escape-hatch shape `NumberFormat` uses on a cell's style:

```xml
<pageSetup orientation="landscape">
  <paperSize kind="a4" />
  <margins left="0.5" right="0.5" top="1" bottom="1" header="0.2" footer="0.2" />
</pageSetup>
```

A macro-enabled workbook's `VbaProject` bytes sit at the workbook level, alongside
`sheets`, not inside any one sheet:

```xml
<workbook>
  <sheets>...</sheets>
  <vbaProject>AQIDBA==</vbaProject>
</workbook>
```

`DefinedNameScope`'s two cases show a different shape than `PaperSize`/`NumberFormat`'s
"kind attribute" trick: `WorkbookScope` carries no data of its own, yet still becomes its
own (empty) element rather than an attribute value, since it sits in a `<choice>` alongside
`SheetScope`, which does carry data:

```xml
<definedNames>
  <definedName name="LocalTotal" formula="Sheet1!$A$2" hidden="true">
    <sheetScope sheetName="Sheet1" />
  </definedName>
  <definedName name="TaxRate" formula="0.075">
    <workbookScope />
  </definedName>
</definedNames>
```

The smaller range-shaped fields (`MergedRange`, `FreezePane`, `AutoFilter`, `ColumnProps`,
`RowProps`) are all straightforward attribute bags or lists of them:

```xml
<mergedRanges>
  <mergedRange topLeft="A1" bottomRight="C1" />
</mergedRanges>
<freezePane rows="1" columns="0" />
<autoFilter topLeft="A1" bottomRight="D11" />
<columnProps>
  <columnProp index="0" width="20" />
</columnProps>
<rowProps>
  <rowProp index="0" height="30" />
</rowProps>
```

`Xml.schemaSet()` loads the compiled schema for validating either direction yourself
(`XDocument.Validate`) - every scenario under `tests/Kookerella.FsOpenXmlDsl.Tests/Examples/`
has a committed `workbook.xml` validated against it this way as part of the same test that
generates it, so the schema and `Xml.fs` itself can never silently drift apart. `Xml.fs`
covers the same worksheet/workbook-level feature set as the rest of this library and the C#
wrapper - cell values, styles, merged ranges, freeze panes, autofilter, column/row sizing,
VBA (base64), defined names, hyperlinks, comments, sheet/workbook protection, print
settings, images (base64), Excel Tables, sparklines, charts, pivot tables (the description
only - loading one doesn't re-run its aggregation, unlike everything else here), conditional
formatting, and data validation.

`Kookerella.FsOpenXmlDsl.Mcp` exposes both directions without writing any F# at all:
`generate_xml`/`create_workbook_from_xml` MCP tools for an AI agent, and `fsopenxmldsl-mcp
convert --lang xml`/`build` CLI commands for anyone else - see that project's own README.

## JSON

`Json.toWorkbook`/`Json.ofWorkbook` (in `Json.fs`) are a fourth way in and out of the DSL,
alongside writing F#/C# directly, code generation, and XML: plain JSON, for a caller whose
tooling speaks JSON rather than XML. The same two concrete uses XML has apply here:

- **Build an `.xlsx` from JSON a transform/generation pipeline already produces** - without
  learning the OOXML schema or this library's own API.
- **Convert an existing `.xlsx` to JSON for version control** - the same determinism
  `Xml.ofWorkbook` has (sorted by cell position, or by name for defined names) applies to
  `Json.ofWorkbook`'s output too, for the same reason: a genuine content change produces a
  small, isolated diff rather than a spurious one from lists getting reshuffled between runs.

```fsharp
open System.Text.Json.Nodes

// JSON -> Workbook -> .xlsx
let wb = JsonNode.Parse(File.ReadAllText "report.json").AsObject() |> Json.toWorkbook
Workbook.save "report.xlsx" wb

// .xlsx -> Workbook -> JSON
let json = Workbook.load "report.xlsx" |> Json.ofWorkbook
File.WriteAllText("report.json", json.ToJsonString())
```

A discriminated union case becomes a single-key JSON object named after the case
(camelCased) when it carries data of its own, or a bare JSON string (also camelCased) when
it's one of several parameterless alternatives - e.g. a cell's value:

```json
{
  "ref": "B2",
  "number": 42.5,
  "style": { "numberFormat": "currency" }
}
```

The same `DataValidation` example as above, in JSON - unlike the XML surface, which
flattens `ValidationAlert`'s fields onto `<dataValidation>` itself, JSON nests both `kind`
and `alert` as their own objects, the more natural shape for this format:

```json
{
  "topLeft": "A2",
  "bottomRight": "A2",
  "kind": { "wholeNumberValidation": { "operator": "greaterThan", "formula1": "0" } },
  "alert": {
    "errorTitle": "Invalid quantity",
    "errorMessage": "Quantity must be a positive whole number."
  }
}
```

The same `ConditionalFormat` example as above, in JSON - `rule` nests one of the seven
cases the same way `kind` does above, and (unlike XML's bare `<fill>`) `fill` always wraps
its `color` under an explicit key:

```json
{
  "topLeft": "A1",
  "bottomRight": "A3",
  "rule": {
    "cellValueRule": {
      "operator": "greaterThan",
      "formula1": "100",
      "style": { "fill": { "color": { "rgb": { "r": 255, "g": 199, "b": 206 } } } }
    }
  }
}
```

The same `Chart` example as above, in JSON - `series` is a plain array, with no need for
the wrapper-vs-item-name trick `<series>`/`<s>` exist for in XML, since a JSON array is
self-delimiting:

```json
{
  "type": "column",
  "title": "Sales by Quarter",
  "showLegend": true,
  "anchorTopLeft": "E1",
  "anchorBottomRight": "L15",
  "categories": { "topLeft": "A2", "bottomRight": "A4" },
  "series": [
    { "name": "B1", "valuesTopLeft": "B2", "valuesBottomRight": "B4" },
    { "name": "C1", "valuesTopLeft": "C2", "valuesBottomRight": "C4" }
  ]
}
```

The same `Table` example as above, in JSON - `columns` is just another plain array, same
as `series`:

```json
{
  "topLeft": "A1",
  "bottomRight": "B4",
  "name": "Calc",
  "columns": [
    { "name": "Qty" },
    { "name": "Doubled", "calculatedFormula": "Calc[Qty]*2" }
  ],
  "style": {
    "name": "TableStyleLight9",
    "showFirstColumn": true,
    "showLastColumn": true,
    "showColumnStripes": true
  }
}
```

The same `SparklineGroup` example as above, in JSON - `color` sits as a plain nested key
alongside the style's other fields, the same way `fill`'s does under `CellStyle`:

```json
{
  "style": {
    "type": "column",
    "lineWeight": 1.5,
    "showNegative": true,
    "color": { "rgb": { "r": 0, "g": 112, "b": 192 } }
  },
  "sparklines": [
    { "cell": "E1", "dataTopLeft": "A1", "dataBottomRight": "D1" }
  ]
}
```

The same `PivotTable` example as above, in JSON - a flat object either way, since there's
nothing here that's a list or a nested structure:

```json
{
  "sourceSheet": "Data",
  "sourceTopLeft": "A1",
  "sourceBottomRight": "C9",
  "rowField": "Region",
  "columnField": "Quarter",
  "valueField": "Sales",
  "aggregation": "average",
  "valueCaption": "Avg Sales",
  "anchorTopLeft": "F1"
}
```

An `Image`'s bytes are a base64 string value, same as `vbaProject` below:

```json
{ "format": "gif", "topLeft": "A1", "bottomRight": "D6", "data": "R0lGODlhAQABAIAAAAAAAP..." }
```

A `Hyperlink`, in JSON:

```json
{
  "topLeft": "A1",
  "bottomRight": "A1",
  "target": { "externalHyperlink": "https://example.com" },
  "tooltip": "Visit site"
}
```
```json
{
  "topLeft": "A2",
  "bottomRight": "B3",
  "target": { "internalHyperlink": "Sheet1!A1" },
  "display": "Go to top"
}
```

A `Comment` - `author` is a plain optional field, omitted rather than an empty string:

```json
{ "cell": "A1", "author": "Alex", "text": "Check this figure" }
{ "cell": "A2", "text": "Unnamed author" }
```

Sheet and workbook protection, in JSON - flat objects, same as the XML:

```json
{ "password": "hunter2", "sheet": true, "formatCells": true, "sort": true, "autoFilter": true }
```

```json
{ "sheets": [ { "name": "Sheet1" } ], "protection": { "password": "hunter2", "lockStructure": true } }
```

`PageSetup` - `PaperSize`'s named cases are a bare string, the mixed-DU convention
`NumberFormat` also uses:

```json
{
  "orientation": "landscape",
  "paperSize": "a4",
  "margins": { "left": 0.5, "right": 0.5, "top": 1, "bottom": 1, "header": 0.2, "footer": 0.2 }
}
```

`VbaProject`, at the workbook level:

```json
{ "sheets": [ { "name": "Sheet1" } ], "vbaProject": "AQIDBA==" }
```

`DefinedNameScope`'s two cases follow the standard JSON convention cleanly, unlike XML's
`<workbookScope />`/`<sheetScope>` split - `WorkbookScope` is simply the bare string, same
treatment as any other parameterless case:

```json
{
  "definedNames": [
    { "name": "LocalTotal", "formula": "Sheet1!$A$2", "scope": { "sheetScope": "Sheet1" }, "hidden": true },
    { "name": "TaxRate", "formula": "0.075", "scope": "workbookScope" }
  ]
}
```

The smaller range-shaped fields, in JSON:

```json
{
  "mergedRanges": [ { "topLeft": "A1", "bottomRight": "C1" } ],
  "freezePane": { "rows": 1, "columns": 0 },
  "autoFilter": { "topLeft": "A1", "bottomRight": "D11" },
  "columnProps": [ { "index": 0, "width": 20 } ],
  "rowProps": [ { "index": 0, "height": 30 } ]
}
```

Unlike XML, .NET has no built-in JSON Schema validator the way `System.Xml.Schema` exists
for XML, so `Json.schema.json` (in the repo, matching this shape) is validated only from
this repo's own test suite (via a test-only `JsonSchema.Net` dependency) rather than exposed
as a public `Json.schemaSet()`-style API. Every scenario under
`tests/Kookerella.FsOpenXmlDsl.Tests/Examples/` has a committed `workbook.json` validated
against it this way too, the same as `workbook.xml` is against `Xml.xsd`, so the schema and
`Json.fs` itself can never silently drift apart there either. `Json.fs` covers the same
worksheet/workbook-level feature set `Xml.fs` does - cell values, styles, merged ranges,
freeze panes, autofilter, column/row sizing, VBA (base64), defined names, hyperlinks,
comments, sheet/workbook protection, print settings, images (base64), Excel Tables,
sparklines, charts, pivot tables (the description only - loading one doesn't re-run its
aggregation, unlike everything else here), conditional formatting, and data validation.

`Kookerella.FsOpenXmlDsl.Mcp` exposes both directions without writing any F# at all:
`generate_json`/`create_workbook_from_json` MCP tools for an AI agent, and `fsopenxmldsl-mcp
convert --lang json`/`build` CLI commands for anyone else - see that project's own README.

## Building and testing

```bash
dotnet build
dotnet test --filter "Category!=Slow"
dotnet run --project samples/Kookerella.FsOpenXmlDsl.Sample
```

The default loop above skips the slow `Category=Slow` tests, which actually invoke
`dotnet fsi` on every generated `Examples/*/script.fsx` (multi-second process startup
each, so ~30-60s total) rather than just checking the generated source parses. Run those
explicitly, after the fast suite has populated the `.fsx` files at least once:

```bash
dotnet test --filter "Category=Slow"
```

Plain `dotnet test` (no filter) runs both groups.
