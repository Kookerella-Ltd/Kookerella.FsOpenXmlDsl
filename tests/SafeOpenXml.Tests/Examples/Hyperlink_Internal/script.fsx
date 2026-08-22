#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Go to Sheet2") ]; hyperlink ((CellRef.ofA1 "A1"), InternalHyperlink "Sheet2!A1") ]
let sheet1 = sheet "Sheet2" [ row [ cell (Text "You made it!") ] ]

let wb = workbook [ sheet0; sheet1 ]

wb |> Workbook.save "output.xlsx"
