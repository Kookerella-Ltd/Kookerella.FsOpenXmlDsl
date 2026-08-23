#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Team"); cell (Text "Score") ]; row [ cell (Text "Alpha"); cell (Number 42.0) ]; row [ cell (Text "Beta"); cell (Number 37.0) ]; EmbeddedChart { Type = ChartBar; Title = None; CategoriesTopLeft = (CellRef.ofA1 "A2"); CategoriesBottomRight = (CellRef.ofA1 "A3"); Series = [ { Name = (CellRef.ofA1 "B1"); ValuesTopLeft = (CellRef.ofA1 "B2"); ValuesBottomRight = (CellRef.ofA1 "B3") } ]; ShowLegend = false; TopLeftAnchor = (CellRef.ofA1 "D1"); BottomRightAnchor = (CellRef.ofA1 "K12") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
