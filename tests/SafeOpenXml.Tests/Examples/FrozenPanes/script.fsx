#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Name", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }); cell (Text "Amount", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }) ]; row [ cell (Text "Row 1"); cell (Number 1.0) ]; row [ cell (Text "Row 2"); cell (Number 2.0) ]; Freeze(1, 0) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
