#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Quarterly Report").WithStyle(CellStyle.Default.AsBold())),
        Row.Of(Cell.Text("Q1"), Cell.Text("Q2"), Cell.Text("Q3"), Cell.Text("Q4")))
    .WithMergedRanges(MergedRange.Of("A1", "D1"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
