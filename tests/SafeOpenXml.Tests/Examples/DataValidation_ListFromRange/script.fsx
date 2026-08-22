#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Small"); cell (Text "Medium"); cell (Text "Large") ]; row [ cell (Text "Size") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), ListFromRangeValidation((CellRef.ofA1 "A1"), (CellRef.ofA1 "C1"))) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
