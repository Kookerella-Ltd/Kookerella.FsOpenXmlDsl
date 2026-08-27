#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Boxed", style = { CellStyle.Default with Border = (Some ({ BorderStyle.None with Left = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })); Right = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })); Top = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })); Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
