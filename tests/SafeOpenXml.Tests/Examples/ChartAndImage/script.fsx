#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Month"); cell (Text "Units") ]; row [ cell (Text "Jan"); cell (Number 10.0) ]; row [ cell (Text "Feb"); cell (Number 14.0) ]; EmbeddedChart { Type = ChartColumn; Title = None; CategoriesTopLeft = (CellRef.ofA1 "A2"); CategoriesBottomRight = (CellRef.ofA1 "A3"); Series = [ { Name = (CellRef.ofA1 "B1"); ValuesTopLeft = (CellRef.ofA1 "B2"); ValuesBottomRight = (CellRef.ofA1 "B3") } ]; ShowLegend = false; TopLeftAnchor = (CellRef.ofA1 "D1"); BottomRightAnchor = (CellRef.ofA1 "H8") }; EmbeddedImage { Data = System.Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7"); Format = Gif; TopLeftAnchor = (CellRef.ofA1 "D10"); BottomRightAnchor = (CellRef.ofA1 "F15") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
