#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Revenue"); cell (Number 1250.0) ]; row [ cell (Text "Costs"); cell (Number 900.0) ]; comment ((CellRef.ofA1 "B1"), "Figure is provisional pending audit.", author = "Alex"); comment ((CellRef.ofA1 "B2"), "Includes one-off relocation costs.", author = "Alex"); comment ((CellRef.ofA1 "A1"), "Double check this label.") ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
