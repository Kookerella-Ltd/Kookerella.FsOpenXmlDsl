#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Sheets can't be added, removed, or renamed") ] ]

let wb = workbook [ sheet0 ] |> withProtection ({ WorkbookProtection.Default with Password = (Some ("hunter2")); LockStructure = (Some (true)) })

wb |> Workbook.save "output.xlsx"
