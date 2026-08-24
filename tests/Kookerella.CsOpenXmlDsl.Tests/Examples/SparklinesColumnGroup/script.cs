#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Number(-2.0), Cell.Number(4.0), Cell.Number(-1.0), Cell.Number(3.0)))
    .WithSparklineGroups(new SparklineGroupEntry(SparklineCell.Of("E1", "A1", "D1")).WithStyle(SparklineStyle.Default.WithType(SparklineType.Column).WithColor(new RgbColor(0, 112, 192)).WithNegative()));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
