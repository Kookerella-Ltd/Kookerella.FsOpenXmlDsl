#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Short code (<= 10 chars)") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), TextLengthValidation(LessThanOrEqual, "10", None)) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
