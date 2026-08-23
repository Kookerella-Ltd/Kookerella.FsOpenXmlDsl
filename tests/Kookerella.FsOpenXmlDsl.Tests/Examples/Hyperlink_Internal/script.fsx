#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Go to Sheet2") ]; hyperlink ((CellRef.ofA1 "A1"), InternalHyperlink "Sheet2!A1") ]
let sheet1 = sheet "Sheet2" [ row [ cell (Text "You made it!") ] ]

let wb = workbook [ sheet0; sheet1 ]

wb |> Workbook.save "output.xlsx"
