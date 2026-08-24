#:project ../../../../src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj

using Kookerella.CsOpenXmlDsl;

var sheet0 = Sheet.Create(
        "Sheet1",
        Row.Of(Cell.Text("Revenue"), Cell.Number(1250.0)),
        Row.Of(Cell.Text("Costs"), Cell.Number(900.0)))
    .WithComments(CommentEntry.Of("B1", "Figure is provisional pending audit.", "Alex"), CommentEntry.Of("B2", "Includes one-off relocation costs.", "Alex"), CommentEntry.Of("A1", "Double check this label."));

var workbook = Workbook.Create(sheet0);

WorkbookIO.Save(workbook, "output.xlsx");
