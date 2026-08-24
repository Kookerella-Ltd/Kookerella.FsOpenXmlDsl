#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Fits to one page wide")))
    .WithPageSetup(PageSetup.Default.WithScaling(new PrintScaling.FitToPage(1, 0)));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
