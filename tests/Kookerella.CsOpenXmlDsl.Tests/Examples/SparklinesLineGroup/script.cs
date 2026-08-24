#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Widgets"), Cell.Number(3.0), Cell.Number(8.0), Cell.Number(5.0), Cell.Number(9.0)),
        Row.Of(Cell.Text("Gadgets"), Cell.Number(6.0), Cell.Number(4.0), Cell.Number(7.0), Cell.Number(2.0)))
    .WithSparklineGroups(new SparklineGroupEntry(SparklineCell.Of("F1", "B1", "E1"), SparklineCell.Of("F2", "B2", "E2")).WithStyle(SparklineStyle.Default.WithHigh().WithLow()));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
