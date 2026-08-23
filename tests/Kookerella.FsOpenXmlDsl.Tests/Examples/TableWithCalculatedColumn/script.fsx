#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Quantity"); cell (Text "Unit Price"); cell (Text "Total") ]; row [ cell (Number 12.0); cell (Number 2.5); cell (Formula("[@Quantity]*[@[Unit Price]]", (Some (30.0)))) ]; row [ cell (Number 5.0); cell (Number 9.0); cell (Formula("[@Quantity]*[@[Unit Price]]", (Some (45.0)))) ]; Table { TopLeft = (CellRef.ofA1 "A1"); BottomRight = (CellRef.ofA1 "C3"); Name = "Orders"; Columns = [ { Name = "Quantity"; CalculatedFormula = None }; { Name = "Unit Price"; CalculatedFormula = None }; { Name = "Total"; CalculatedFormula = (Some ("[@Quantity]*[@[Unit Price]]")) } ]; Style = { TableStyle.Default with Name = (Some ("TableStyleLight9")); ShowRowStripes = false; ShowColumnStripes = true } } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
