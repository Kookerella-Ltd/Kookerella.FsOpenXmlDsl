#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Protected sheet") ]; Protect({ SheetProtection.Default with Password = (Some ("hunter2")); FormatCells = (Some (true)); Sort = (Some (true)); AutoFilter = (Some (true)) }) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
