#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Quantity")))
    .WithDataValidations(DataValidationEntry.Of("A2", "A2", new ValidationKind.WholeNumberValidation(ComparisonOperator.GreaterThan, "0", null)).WithAlert(ValidationAlert.Default.WithErrorTitle("Invalid quantity").WithErrorMessage("Quantity must be a positive whole number.")));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
