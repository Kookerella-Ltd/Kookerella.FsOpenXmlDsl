#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "First",
        Row.Of(Cell.Text("a")));

var sheet1 = Sheet.Create(
        "Second",
        Row.Of(Cell.Text("b")));

var workbook = Workbook.Create(sheet0, sheet1);

WorkbookIO.Save(workbook, "output.xlsx");
