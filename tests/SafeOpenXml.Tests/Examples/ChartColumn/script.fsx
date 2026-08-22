#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Quarter"); cell (Text "North"); cell (Text "South") ]; row [ cell (Text "Q1"); cell (Number 12.0); cell (Number 9.0) ]; row [ cell (Text "Q2"); cell (Number 15.0); cell (Number 11.0) ]; row [ cell (Text "Q3"); cell (Number 9.0); cell (Number 14.0) ]; EmbeddedChart { Type = ChartColumn; Title = (Some ("Sales by Quarter")); CategoriesTopLeft = (CellRef.ofA1 "A2"); CategoriesBottomRight = (CellRef.ofA1 "A4"); Series = [ { Name = (CellRef.ofA1 "B1"); ValuesTopLeft = (CellRef.ofA1 "B2"); ValuesBottomRight = (CellRef.ofA1 "B4") }; { Name = (CellRef.ofA1 "C1"); ValuesTopLeft = (CellRef.ofA1 "C2"); ValuesBottomRight = (CellRef.ofA1 "C4") } ]; ShowLegend = true; TopLeftAnchor = (CellRef.ofA1 "E1"); BottomRightAnchor = (CellRef.ofA1 "L15") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
