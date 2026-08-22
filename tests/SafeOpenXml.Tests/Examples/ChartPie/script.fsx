#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Segment"); cell (Text "Share") ]; row [ cell (Text "Retail"); cell (Number 55.0) ]; row [ cell (Text "Wholesale"); cell (Number 30.0) ]; row [ cell (Text "Online"); cell (Number 15.0) ]; EmbeddedChart { Type = ChartPie; Title = (Some ("Revenue Share")); CategoriesTopLeft = (CellRef.ofA1 "A2"); CategoriesBottomRight = (CellRef.ofA1 "A4"); Series = [ { Name = (CellRef.ofA1 "B1"); ValuesTopLeft = (CellRef.ofA1 "B2"); ValuesBottomRight = (CellRef.ofA1 "B4") } ]; ShowLegend = true; TopLeftAnchor = (CellRef.ofA1 "D1"); BottomRightAnchor = (CellRef.ofA1 "K14") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
