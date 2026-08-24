#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Wide report")))
    .WithPageSetup(PageSetup.Default.WithOrientation(PageOrientation.Landscape).WithPaperSize(new PaperSize.A4()).WithScaling(new PrintScaling.ScalePercent(85)).WithMargins(PageMargins.Default.WithLeft(0.5).WithRight(0.5)).WithHeader("&C&\"Arial,Bold\"Quarterly Report").WithFooter("&LPage &P of &N&R&D"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
