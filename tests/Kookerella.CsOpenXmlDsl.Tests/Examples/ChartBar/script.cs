#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Team"), Cell.Text("Score")),
        Row.Of(Cell.Text("Alpha"), Cell.Number(42.0)),
        Row.Of(Cell.Text("Beta"), Cell.Number(37.0)))
    .WithCharts(ChartEntry.Of(ChartType.Bar, "A2", "A3", "D1", "K12", ChartSeries.Of("B1", "B2", "B3")));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
