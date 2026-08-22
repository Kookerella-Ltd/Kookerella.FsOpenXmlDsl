#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number 10.0); cell (Number 20.0) ]; row [ cell (Number 30.0); cell (Number 5.0) ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A2"), FormulaRule("A1>B1", { CellStyle.Default with Fill = (Some ({ Color = Rgb(198uy, 239uy, 206uy) })) })) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
