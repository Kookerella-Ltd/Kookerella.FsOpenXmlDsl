#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Data",
        Row.Of(Cell.Text("Category"), Cell.Text("Amount")),
        Row.Of(Cell.Text("A"), Cell.Number(3.0)),
        Row.Of(Cell.Text("B"), Cell.Number(7.0)),
        Row.Of(Cell.Text("A"), Cell.Number(4.0)));

var sheet1 = Sheet.Create(
        "Report",
        Row.Of(Cell.Text("Pivot table below:")))
    .WithPivotTables(PivotTableEntry.Of("A1", "B4", "Category", "Amount", "A3").WithSourceSheet("Data").WithAggregation(PivotAggregation.Count));

var workbook = Workbook.Create(sheet0, sheet1);

WorkbookIO.Save(workbook, "output.xlsx");
