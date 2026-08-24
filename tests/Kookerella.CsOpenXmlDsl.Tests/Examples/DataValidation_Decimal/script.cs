#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Fraction (0-1)")))
    .WithDataValidations(DataValidationEntry.Of("A2", "A2", new ValidationKind.DecimalValidation(ComparisonOperator.Between, "0", "1")));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
