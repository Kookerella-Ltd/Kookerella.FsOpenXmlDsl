#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Region"), Cell.Text("Quarter"), Cell.Text("Sales")),
        Row.Of(Cell.Text("East"), Cell.Text("Q1"), Cell.Number(10.0)),
        Row.Of(Cell.Text("East"), Cell.Text("Q2"), Cell.Number(5.0)),
        Row.Of(Cell.Text("West"), Cell.Text("Q1"), Cell.Number(20.0)),
        Row.Of(Cell.Text("West"), Cell.Text("Q2"), Cell.Number(15.0)))
    .WithPivotTables(PivotTableEntry.Of("A1", "C5", "Region", "Sales", "E1").WithColumnField("Quarter").WithValueCaption("Total Sales"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
