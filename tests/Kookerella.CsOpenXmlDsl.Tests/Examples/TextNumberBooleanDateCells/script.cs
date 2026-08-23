#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Item"), Cell.Text("Qty"), Cell.Text("Price"), Cell.Text("Total")),
        Row.Of(Cell.Text("Widgets"), Cell.Number(4.0), Cell.Number(2.5), Cell.Formula("B2*C2", 10.0)),
        Row.Of(Cell.Boolean(true), Cell.Date(new DateTime(639028224000000000L, DateTimeKind.Unspecified))));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
