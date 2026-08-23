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

## Scope

This is a deliberately narrow first pass, not the whole F# library ported to C#: cell
values (text/number/boolean/date/formula), basic styling (font/fill/border/alignment/
number format), merged ranges, freeze panes, autofilter, and `Save`/`Load`. Tables, charts,
images, pivot tables, sparklines, VBA, conditional formatting, data validation, hyperlinks,
comments, print settings, defined names, protection, and code generation aren't exposed
here - reference `Kookerella.FsOpenXmlDsl` directly for those (this wrapper doesn't stop
you from mixing both in the same project).

Formula cells never carry a cached value from this API beyond what you explicitly pass to
`Cell.Formula` - see the main project's README for why that matters for anything that isn't
opened in real Excel first (there's no formula evaluation engine anywhere in this stack).
