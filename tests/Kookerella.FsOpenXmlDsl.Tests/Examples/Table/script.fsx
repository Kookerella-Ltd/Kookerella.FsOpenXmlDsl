#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Item"); cell (Text "Quantity") ]; row [ cell (Text "Widgets"); cell (Number 12.0) ]; row [ cell (Text "Gadgets"); cell (Number 5.0) ]; Table { TopLeft = (CellRef.ofA1 "A1"); BottomRight = (CellRef.ofA1 "B3"); Name = "Inventory"; Columns = [ { Name = "Item"; CalculatedFormula = None }; { Name = "Quantity"; CalculatedFormula = None } ]; Style = TableStyle.Default } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
