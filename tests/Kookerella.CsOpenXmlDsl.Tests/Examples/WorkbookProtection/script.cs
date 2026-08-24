#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Sheets can't be added, removed, or renamed")));

var workbook = Workbook.Create(sheet0)
    .WithProtection(WorkbookProtection.Default.WithPassword("hunter2").WithLockStructure(true));

WorkbookIO.Save(workbook, "output.xlsx");
