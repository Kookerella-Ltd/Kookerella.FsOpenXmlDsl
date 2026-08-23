#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Quantity"), Cell.Text("Unit Price"), Cell.Text("Total")),
        Row.Of(Cell.Number(12.0), Cell.Number(2.5), Cell.Formula("[@Quantity]*[@[Unit Price]]", 30.0)),
        Row.Of(Cell.Number(5.0), Cell.Number(9.0), Cell.Formula("[@Quantity]*[@[Unit Price]]", 45.0)))
    .WithTables(TableEntry.Of("A1", "C3", "Orders", new TableColumn("Quantity"), new TableColumn("Unit Price"), new TableColumn("Total", "[@Quantity]*[@[Unit Price]]")).WithStyle(TableStyle.Default.WithName("TableStyleLight9").WithoutRowStripes().WithColumnStripes()));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
