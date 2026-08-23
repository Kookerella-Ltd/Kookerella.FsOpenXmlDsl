#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Small"); cell (Text "Medium"); cell (Text "Large") ]; row [ cell (Text "Size") ]; dataValidation ((CellRef.ofA1 "A2"), (CellRef.ofA1 "A2"), ListFromRangeValidation((CellRef.ofA1 "A1"), (CellRef.ofA1 "C1"))) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
