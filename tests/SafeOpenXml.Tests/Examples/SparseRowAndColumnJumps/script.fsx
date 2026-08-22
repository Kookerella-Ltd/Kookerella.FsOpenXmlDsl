#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "A1") ]; row ([ cell (Text "A3"); cell (Text "C4", col = 2) ], index = 2) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
