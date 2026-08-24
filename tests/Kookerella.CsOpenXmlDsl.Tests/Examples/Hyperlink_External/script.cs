#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Open-XML-SDK on GitHub")))
    .WithHyperlinks(HyperlinkEntry.Of("A1", new HyperlinkTarget.ExternalHyperlink("https://github.com/dotnet/Open-XML-SDK")).WithTooltip("Open in browser").WithDisplay("dotnet/Open-XML-SDK"));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
