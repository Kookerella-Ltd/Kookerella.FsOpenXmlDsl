#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Enter quantity:"); cell (Number 0.0, style = { CellStyle.Default with Protection = (Some ({ CellProtection.Default with Locked = false })) }) ]; Protect(SheetProtection.Default) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
