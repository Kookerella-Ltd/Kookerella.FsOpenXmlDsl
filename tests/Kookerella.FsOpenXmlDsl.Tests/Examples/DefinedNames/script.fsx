#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number 0.075) ]; row [ cell (Number 100.0) ]; row [ cell (Formula("B1*(1+TaxRate)", (Some (107.5)))) ] ]

let wb = workbook [ sheet0 ] |> withDefinedNames [ definedName "TaxRate" "Sheet1!$A$1"; sheetScopedDefinedName "Sheet1" "LocalTotal" "Sheet1!$A$2" ]

wb |> Workbook.save "output.xlsx"
