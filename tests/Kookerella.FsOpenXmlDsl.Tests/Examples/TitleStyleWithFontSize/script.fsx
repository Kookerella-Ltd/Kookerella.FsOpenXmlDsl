#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "INVOICE", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Size = (Some (16.0)); Bold = true; Italic = true; Underline = true; Color = (Some (Rgb(68uy, 84uy, 106uy))) })) }) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
