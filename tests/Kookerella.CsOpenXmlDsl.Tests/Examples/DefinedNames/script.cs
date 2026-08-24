#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Number(0.075)),
        Row.Of(Cell.Number(100.0)),
        Row.Of(Cell.Formula("B1*(1+TaxRate)", 107.5)));

var workbook = Workbook.Create(sheet0)
    .WithDefinedNames(DefinedNameEntry.Of("TaxRate", "Sheet1!$A$1"), DefinedNameEntry.Of("LocalTotal", "Sheet1!$A$2", "Sheet1"));

WorkbookIO.Save(workbook, "output.xlsx");
