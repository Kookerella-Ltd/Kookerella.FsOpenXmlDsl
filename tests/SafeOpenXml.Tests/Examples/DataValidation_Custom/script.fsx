#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Must be a number") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), CustomValidation("ISNUMBER(A2)"), allowBlank = false, inputTitle = "Note", inputMessage = "Enter a numeric value.") ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
