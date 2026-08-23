#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Widgets"); cell (Number 4.0); cell (Number 2.5); cell (Formula("B1*C1", (Some (10.0)))) ]; row [ cell (Text "Gadgets"); cell (Number 2.0); cell (Number 19.99); cell (Formula("B2*C2", (Some (39.98)))) ]; row [ cell (Text "Total"); cell (Formula("SUM(D1:D2)", (Some (49.98)))) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
