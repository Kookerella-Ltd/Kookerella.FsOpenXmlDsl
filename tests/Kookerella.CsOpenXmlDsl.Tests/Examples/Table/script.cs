#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Item"), Cell.Text("Quantity")),
        Row.Of(Cell.Text("Widgets"), Cell.Number(12.0)),
        Row.Of(Cell.Text("Gadgets"), Cell.Number(5.0)))
    .WithTables(TableEntry.Of("A1", "B3", "Inventory", new TableColumn("Item"), new TableColumn("Quantity")));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
