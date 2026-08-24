#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Number(10.0)),
        Row.Of(Cell.Number(50.0)),
        Row.Of(Cell.Number(90.0)))
    .WithConditionalFormats(ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.DataBarRule(RgbColor.Blue)));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
