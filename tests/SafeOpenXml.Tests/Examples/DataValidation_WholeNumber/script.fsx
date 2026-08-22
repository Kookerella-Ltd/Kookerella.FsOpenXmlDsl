#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Quantity") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), WholeNumberValidation(GreaterThan, "0", None), errorTitle = "Invalid quantity", errorMessage = "Quantity must be a positive whole number.") ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
