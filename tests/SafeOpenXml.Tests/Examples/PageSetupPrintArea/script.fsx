#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Included") ]; row [ cell (Text "Excluded") ]; row [ cell (Text "Also included") ]; PageSetup({ PageSetup.Default with PrintArea = [ ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A1")); ((CellRef.ofA1 "A3"), (CellRef.ofA1 "A3")) ] }) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
