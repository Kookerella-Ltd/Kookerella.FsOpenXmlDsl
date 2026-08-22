#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Quarterly Report", style = { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true; Color = (Some (Rgb(255uy, 255uy, 255uy))) })); Fill = (Some ({ Color = Rgb(68uy, 84uy, 106uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })) }) ]; row [ cell (Text "Q1"); cell (Text "Q2"); cell (Text "Q3"); cell (Text "Q4") ]; Merge((CellRef.ofA1 "A1"), (CellRef.ofA1 "D1")) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
