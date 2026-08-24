#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Apple")),
        Row.Of(Cell.Text("Banana")),
        Row.Of(Cell.Text("Apple")))
    .WithConditionalFormats(ConditionalFormatEntry.Of("A1", "A3", new ConditionalFormatRule.UniqueValuesRule(CellStyle.Default.WithFillColor(new RgbColor(198, 239, 206)))));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
