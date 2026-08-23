#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number 10.0) ]; row [ cell (Number 50.0) ]; row [ cell (Number 90.0) ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A3"), DataBarRule(Rgb(0uy, 0uy, 255uy))) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
