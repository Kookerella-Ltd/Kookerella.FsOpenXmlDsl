#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Included") ]; row [ cell (Text "Excluded") ]; row [ cell (Text "Also included") ]; PageSetup({ PageSetup.Default with PrintArea = [ ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A1")); ((CellRef.ofA1 "A3"), (CellRef.ofA1 "A3")) ] }) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
