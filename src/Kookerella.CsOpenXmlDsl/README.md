# Kookerella.CsOpenXmlDsl

An idiomatic, immutable, fluent C# wrapper over
[Kookerella.FsOpenXmlDsl](https://www.nuget.org/packages/Kookerella.FsOpenXmlDsl) - build
and read Excel workbooks (.xlsx/.xlsm) from C# without touching F# discriminated unions or
option types directly.

Every type here is an immutable C# `record` - every `With*`/`As*` method returns a new
instance rather than mutating in place, so a style or sheet can be built up once and safely
reused across many cells without aliasing surprises. The only place this library does any
I/O at all is `WorkbookIO` - every other type is pure data.

```csharp
using Kookerella.CsOpenXmlDsl;

var headerStyle = CellStyle.Default.AsBold().WithFillColor(new RgbColor(220, 220, 220));

var sheet = Sheet.Create("Sheet1",
    Row.Of(
        Cell.Text("Item").WithStyle(headerStyle),
        Cell.Text("Amount").WithStyle(headerStyle)),
    Row.Of(
        Cell.Text("Widgets"),
        Cell.Number(42.5)));

WorkbookIO.Save(Workbook.Create(sheet), "out.xlsx");

var roundTripped = WorkbookIO.Load("out.xlsx");
```

Merged ranges, a frozen header row, and an autofilter range are sheet-level facts, set
fluently the same way:

```csharp
var sheet = Sheet.Create("Sheet1", /* ...rows... */)
    .WithMergedRanges(MergedRange.Of("A1", "D1"))
    .WithFreezePane(1, 0) // freeze the header row
    .WithAutoFilter(AutoFilterRange.Of("A2", "D10"));
```

`CellPosition` addresses cells zero-based (`Row 0, Column 0` is "A1") and converts to/from
A1-style strings via `CellPosition.FromA1("B3")` / `position.ToA1()`.

A VBA project (macros) is opaque bytes, same treatment as the F# core - nothing in this
stack parses, generates, or edits VBA source, it only embeds and hands back exactly what
you give it:

```csharp
var workbook = Workbook.Create(sheet)
    .WithVbaProject(File.ReadAllBytes("vbaProject.bin"));

WorkbookIO.Save(workbook, "out.xlsm"); // .xlsm, not .xlsx - see below
```

Save to an `.xlsm` path once a VBA project is attached - the file's content type switches
to macro-enabled automatically, but real Excel also expects the extension to match before
it will trust and run macros regardless of what the content type says.

Excel Tables (`ListObject`s, the things structured references like `Table1[Column]` point
at) are added the same fluent way, with columns and a visual style:

```csharp
var sheet = Sheet.Create("Sheet1",
        Row.Of(Cell.Text("Item"), Cell.Text("Quantity")),
        Row.Of(Cell.Text("Widgets"), Cell.Number(12)))
    .WithTables(
        TableEntry.Of("A1", "B2", "Inventory", new TableColumn("Item"), new TableColumn("Quantity"))
            .WithStyle(TableStyle.Default.WithName("TableStyleLight9")));
```

This wrapper doesn't synthesize the header row's text for you - it must already be there as
ordinary cells, the same way merged ranges/autofilter only describe metadata layered on top
of cells you've already placed. `Columns`' count must equal the range's width and every
column name must be unique (genuine Excel/OOXML requirements) - `WorkbookIO.Save` throws an
`ArgumentException` if either is violated rather than silently producing a file Excel would
refuse to open cleanly.

Charts are anchored over a range of cells (a "move and size with cells" anchor, matching how
tables/merged ranges are already addressed) rather than a pixel-precise floating position:

```csharp
var sheet = Sheet.Create("Sheet1",
        Row.Of(Cell.Text("Quarter"), Cell.Text("North"), Cell.Text("South")),
        Row.Of(Cell.Text("Q1"), Cell.Number(12), Cell.Number(9)),
        Row.Of(Cell.Text("Q2"), Cell.Number(15), Cell.Number(11)))
    .AddChart(
        ChartEntry
            .Of(ChartType.Column, "A2", "A3", "E1", "L15",
                ChartSeries.Of("B1", "B2", "B3"),
                ChartSeries.Of("C1", "C2", "C3"))
            .WithTitle("Sales by Quarter")
            .WithLegend());
```

A series' name is a reference to the cell that holds it (its header, typically), not a
literal string, matching how a real Excel chart's series name live-updates if that cell's
text changes. `ChartType` covers `Column`/`Bar`/`Line`/`Pie` - the same set the F# core
models, no scatter/area/stock/3-D/stacked variants in either layer.

Images are raster files embedded and anchored the same "move and size with cells" way as
charts and tables - this wrapper does no decoding of its own, `Data` is exactly the bytes
of the image file on disk, handed back unchanged on read:

```csharp
var sheet = Sheet.Create("Sheet1", Row.Of(Cell.Text("Logo below:")))
    .AddImage(ImageEntry.Of(File.ReadAllBytes("logo.png"), ImageFormat.Png, "A3", "C10"));
```

`ImageFormat` covers `Png`/`Jpeg`/`Gif`/`Bmp` - the four formats every Excel version has
supported natively, matching the F# core (TIFF/SVG/EMF/WMF aren't modeled in either layer).

Pivot tables are unlike everything else above: `WorkbookIO.Save` doesn't just describe a
reference for Excel to resolve later, it actually performs the grouping and aggregation
itself, since a pivot table's file format bakes the *computed* result into the workbook in
several places that all have to agree. Scoped to the single most common shape - exactly one
row field, at most one optional column field, and exactly one value field:

```csharp
var sheet = Sheet.Create("Sheet1",
        Row.Of(Cell.Text("Region"), Cell.Text("Sales")),
        Row.Of(Cell.Text("East"), Cell.Number(10)),
        Row.Of(Cell.Text("West"), Cell.Number(20)),
        Row.Of(Cell.Text("East"), Cell.Number(5)))
    .AddPivotTable(
        PivotTableEntry.Of("A1", "B4", "Region", "Sales", "D1")
            .WithAggregation(PivotAggregation.Sum));
```

`RowField`/`ColumnField`/`ValueField` must exactly match header cell text in the source
range's first row (a genuine Excel/OOXML requirement - the source range must already have a
header row, the same way tables do). The source range can live on a different sheet than
the pivot table itself via `WithSourceSheet(name)`. `PivotAggregation` covers
`Sum`/`Count`/`CountNumbers`/`Average`/`Min`/`Max`, defaulting to `Sum`.

Sparklines are grouped so several can be styled together at once - a sheet can have several
independently-styled groups:

```csharp
var sheet = Sheet.Create("Sheet1",
        Row.Of(Cell.Number(-2), Cell.Number(4), Cell.Number(-1), Cell.Number(3)))
    .AddSparklineGroup(
        new SparklineGroupEntry(SparklineCell.Of("E1", "A1", "D1"))
            .WithStyle(
                SparklineStyle.Default
                    .WithType(SparklineType.Column)
                    .WithColor(new RgbColor(0, 112, 192))
                    .WithNegative()));
```

`SparklineType` covers `Line`/`Column`/`WinLoss` (Excel's "Win/Loss" chart). `Color` is the
one color modeled (the main sparkline color; Excel's separate negative-point/axis/marker
colors default to its own automatic choices), and the `With*` toggles (`WithHigh`,
`WithLow`, `WithFirst`, `WithLast`, `WithNegative`, `WithMarkers`) mirror the "highlight
these points" options on Excel's Sparkline Design ribbon - `WithMarkers`/`WithLineWeight`
are only meaningful for `SparklineType.Line`.

Conditional formatting rules are a closed set of cases (`ConditionalFormatRule.CellValueRule`/
`.FormulaRule`/`.ColorScale2`/`.ColorScale3`/`.DataBarRule`/`.DuplicateValuesRule`/
`.UniqueValuesRule`), the same "sealed hierarchy" pattern `CellValue` uses:

```csharp
var sheet = Sheet.Create("Sheet1",
        Row.Of(Cell.Number(50)),
        Row.Of(Cell.Number(150)),
        Row.Of(Cell.Number(90)))
    .AddConditionalFormat(
        ConditionalFormatEntry.Of(
            "A1", "A3",
            new ConditionalFormatRule.CellValueRule(
                ComparisonOperator.GreaterThan, "100", null,
                CellStyle.Default.WithFillColor(new RgbColor(255, 199, 206)))));
```

`CellValueRule`'s `Formula1`/`Formula2` and `FormulaRule`'s `Formula` are raw formula text
(same convention as `Cell.Formula`) - for `CellValueRule` these are literal values or cell
references compared against, not `=`-prefixed formulas. `ComparisonOperator` is shared with
data validation's own numeric rules once those are exposed. Icon sets, "top/bottom N", and
the text/blank/error-contains rule kinds aren't modeled - reference `Kookerella.FsOpenXmlDsl`
directly for those.

`CsCodeGen.Generate` renders a `Workbook` back out as a self-contained C# file that
regenerates an equivalent file when run - the reverse of `WorkbookIO.Load` one level
further: loading turns a file into these types, this turns those types into C# *source
text*. It targets .NET 10's "file-based apps" feature (`dotnet run script.cs` - no
`.csproj` needed), so the emitted file is directly runnable, not just a snippet to paste
into an existing project:

```csharp
var script = CsCodeGen.Generate(
    ["#:package Kookerella.CsOpenXmlDsl@0.1.0"],
    "regenerated.xlsx",
    loadedWorkbook);

File.WriteAllText("regenerate.cs", script);
// then: dotnet run regenerate.cs
```

The first argument is whatever raw `#:package`/`#:project` directive lines the emitted file
needs to locate this assembly - `CsCodeGen` has no opinion on that, since it depends
entirely on where the file ends up living relative to your own build (pass a `#:project
../path/to/Kookerella.CsOpenXmlDsl.csproj` line instead when generating against a local
checkout rather than the published package). Generated code only mentions what isn't
already implied by a type's own defaults (e.g. `CellStyle.Default`, `TableStyle.Default`),
and only emits an explicit `.AtIndex`/`.AtColumn` where a row or cell's position actually
deviates from strict sequential numbering - so it reads close to how a human would write it
by hand.

See [`tests/Kookerella.CsOpenXmlDsl.Tests/Examples/`](../../tests/Kookerella.CsOpenXmlDsl.Tests/Examples/)
for a real, openable `output.xlsx` plus a runnable `script.cs` per feature covered above -
open any single example in Excel, or `cd` into its folder and run `dotnet run script.cs` to
regenerate it yourself.

## Scope

This is a deliberately narrow first pass, not the whole F# library ported to C#: cell
values (text/number/boolean/date/formula), basic styling (font/fill/border/alignment/
number format), merged ranges, freeze panes, autofilter, tables, charts, images, pivot
tables (single row/column/value field only), sparklines, conditional formatting, VBA (as
opaque bytes), `Save`/`Load`, and code generation (`CsCodeGen`, covering everything this
wrapper itself models). Data validation, hyperlinks, comments, print settings, defined
names, and protection aren't exposed here - reference `Kookerella.FsOpenXmlDsl` directly
for those (this wrapper doesn't stop you from mixing both in the same project).

Formula cells never carry a cached value from this API beyond what you explicitly pass to
`Cell.Formula` - see the main project's README for why that matters for anything that isn't
opened in real Excel first (there's no formula evaluation engine anywhere in this stack).
