#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Name", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }); cell (Text "Amount", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }); cell (Text "Purchased", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }) ]; row [ cell (Text "Widgets"); cell (Number 42.5); cell (Date (System.DateTime.FromOADate(46037.0)), style = { CellStyle.Default with NumberFormat = (Some (ShortDate)) }) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
