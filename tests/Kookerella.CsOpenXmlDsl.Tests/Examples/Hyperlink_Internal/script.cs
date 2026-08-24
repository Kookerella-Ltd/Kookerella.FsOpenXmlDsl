#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Go to Sheet2")))
    .WithHyperlinks(HyperlinkEntry.Of("A1", new HyperlinkTarget.InternalHyperlink("Sheet2!A1")));

var sheet1 = Sheet.Create(
        "Sheet2",
        Row.Of(Cell.Text("You made it!")));

var workbook = Workbook.Create(sheet0, sheet1);

WorkbookIO.Save(workbook, "output.xlsx");
