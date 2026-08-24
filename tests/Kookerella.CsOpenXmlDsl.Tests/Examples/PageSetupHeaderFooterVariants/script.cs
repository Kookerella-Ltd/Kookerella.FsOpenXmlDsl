#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Report")))
    .WithPageSetup(PageSetup.Default.WithHeader("&CPage &P").WithEvenFooter("&L&F").WithFirstHeader("&CCover Page"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
