#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Small"), Cell.Text("Medium"), Cell.Text("Large")),
        Row.Of(Cell.Text("Size")))
    .WithDataValidations(DataValidationEntry.Of("A2", "A2", new ValidationKind.ListFromRangeValidation(CellPosition.FromA1("A1"), CellPosition.FromA1("C1"))));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
