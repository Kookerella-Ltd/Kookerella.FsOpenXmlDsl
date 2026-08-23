#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Quarter"), Cell.Text("North"), Cell.Text("South")),
        Row.Of(Cell.Text("Q1"), Cell.Number(12.0), Cell.Number(9.0)),
        Row.Of(Cell.Text("Q2"), Cell.Number(15.0), Cell.Number(11.0)),
        Row.Of(Cell.Text("Q3"), Cell.Number(9.0), Cell.Number(14.0)))
    .WithCharts(ChartEntry.Of(ChartType.Column, "A2", "A4", "E1", "L15", ChartSeries.Of("B1", "B2", "B4"), ChartSeries.Of("C1", "C2", "C4")).WithTitle("Sales by Quarter").WithLegend());

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
