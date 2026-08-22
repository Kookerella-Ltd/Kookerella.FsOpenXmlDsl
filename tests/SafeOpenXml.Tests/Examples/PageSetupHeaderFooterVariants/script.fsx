#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Report") ]; PageSetup({ PageSetup.Default with Header = (Some ("&CPage &P")); EvenFooter = (Some ("&L&F")); FirstHeader = (Some ("&CCover Page")) }) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
