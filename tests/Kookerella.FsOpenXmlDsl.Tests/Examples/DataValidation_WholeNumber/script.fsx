#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Quantity") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), WholeNumberValidation(GreaterThan, "0", None), errorTitle = "Invalid quantity", errorMessage = "Quantity must be a positive whole number.") ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
