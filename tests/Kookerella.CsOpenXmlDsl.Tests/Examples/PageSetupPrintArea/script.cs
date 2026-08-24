#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Included")),
        Row.Of(Cell.Text("Excluded")),
        Row.Of(Cell.Text("Also included")))
    .WithPageSetup(PageSetup.Default.WithPrintArea(("A1", "A1"), ("A3", "A3")));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
