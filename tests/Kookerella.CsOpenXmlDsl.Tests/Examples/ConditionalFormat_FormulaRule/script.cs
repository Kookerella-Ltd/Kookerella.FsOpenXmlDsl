#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Number(10.0), Cell.Number(20.0)),
        Row.Of(Cell.Number(30.0), Cell.Number(5.0)))
    .WithConditionalFormats(ConditionalFormatEntry.Of("A1", "A2", new ConditionalFormatRule.FormulaRule("A1>B1", CellStyle.Default.WithFillColor(new RgbColor(198, 239, 206)))));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
