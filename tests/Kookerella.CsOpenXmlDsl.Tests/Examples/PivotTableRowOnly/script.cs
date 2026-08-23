#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Region"), Cell.Text("Sales")),
        Row.Of(Cell.Text("East"), Cell.Number(10.0)),
        Row.Of(Cell.Text("West"), Cell.Number(20.0)),
        Row.Of(Cell.Text("East"), Cell.Number(5.0)),
        Row.Of(Cell.Text("West"), Cell.Number(15.0)))
    .WithPivotTables(PivotTableEntry.Of("A1", "B5", "Region", "Sales", "D1"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
