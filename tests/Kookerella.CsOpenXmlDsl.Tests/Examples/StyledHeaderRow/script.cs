#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Number(42.5).WithStyle(CellStyle.Default.AsBold().AsItalic().WithFontColor(RgbColor.White).WithFillColor(new RgbColor(68, 84, 106)).WithBorder(CellBorder.None.WithAllSides(new BorderSide(BorderLineStyle.Thin))).WithHorizontalAlignment(HorizontalCellAlignment.Center).AsWrapText().WithNumberFormat(NumberFormatKind.Currency))));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
