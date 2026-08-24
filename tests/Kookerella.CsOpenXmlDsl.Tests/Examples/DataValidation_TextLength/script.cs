#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Short code (<= 10 chars)")))
    .WithDataValidations(DataValidationEntry.Of("A2", "A2", new ValidationKind.TextLengthValidation(ComparisonOperator.LessThanOrEqual, "10", null)));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
