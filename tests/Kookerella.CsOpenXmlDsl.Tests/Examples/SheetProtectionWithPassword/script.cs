#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Protected sheet")))
    .WithProtection(SheetProtection.Default.WithPassword("hunter2").WithFormatCellsBlocked(true).WithSortBlocked(true).WithAutoFilterBlocked(true));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
