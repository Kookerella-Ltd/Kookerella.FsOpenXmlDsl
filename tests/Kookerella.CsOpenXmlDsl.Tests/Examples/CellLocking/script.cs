#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Enter quantity:"), Cell.Number(0.0).WithStyle(CellStyle.Default.WithProtection(CellProtection.Default.WithLocked(false)))))
    .WithProtection(SheetProtection.Default);

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
