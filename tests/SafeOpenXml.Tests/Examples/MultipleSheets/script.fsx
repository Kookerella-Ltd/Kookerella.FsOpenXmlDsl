#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "First" [ row [ cell (Text "one") ] ]
let sheet1 = sheet "Second" [ row [ cell (Number 2.0) ] ]

let wb = workbook [ sheet0; sheet1 ]

wb |> Workbook.save "output.xlsx"
