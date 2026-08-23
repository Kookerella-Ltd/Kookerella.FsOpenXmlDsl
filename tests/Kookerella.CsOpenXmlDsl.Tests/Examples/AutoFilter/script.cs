#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Name").WithStyle(CellStyle.Default.AsBold()), Cell.Text("Amount").WithStyle(CellStyle.Default.AsBold()), Cell.Text("Region").WithStyle(CellStyle.Default.AsBold())),
        Row.Of(Cell.Text("Widgets"), Cell.Number(42.5), Cell.Text("North")),
        Row.Of(Cell.Text("Gadgets"), Cell.Number(19.99), Cell.Text("South")))
    .WithAutoFilter(AutoFilterRange.Of("A1", "C3"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
