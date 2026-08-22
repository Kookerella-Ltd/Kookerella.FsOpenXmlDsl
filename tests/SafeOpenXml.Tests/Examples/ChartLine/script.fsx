#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Month"); cell (Text "2025"); cell (Text "2026") ]; row [ cell (Text "Jan"); cell (Number 100.0); cell (Number 120.0) ]; row [ cell (Text "Feb"); cell (Number 110.0); cell (Number 115.0) ]; row [ cell (Text "Mar"); cell (Number 105.0); cell (Number 130.0) ]; EmbeddedChart { Type = ChartLine; Title = (Some ("Monthly Trend")); CategoriesTopLeft = (CellRef.ofA1 "A2"); CategoriesBottomRight = (CellRef.ofA1 "A4"); Series = [ { Name = (CellRef.ofA1 "B1"); ValuesTopLeft = (CellRef.ofA1 "B2"); ValuesBottomRight = (CellRef.ofA1 "B4") }; { Name = (CellRef.ofA1 "C1"); ValuesTopLeft = (CellRef.ofA1 "C2"); ValuesBottomRight = (CellRef.ofA1 "C4") } ]; ShowLegend = true; TopLeftAnchor = (CellRef.ofA1 "E1"); BottomRightAnchor = (CellRef.ofA1 "L15") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
