#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Name").WithStyle(CellStyle.Default.AsBold()), Cell.Text("Amount").WithStyle(CellStyle.Default.AsBold())),
        Row.Of(Cell.Text("Row 1"), Cell.Number(1.0)),
        Row.Of(Cell.Text("Row 2"), Cell.Number(2.0)))
    .WithFreezePane(new FreezePane(1, 0));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
