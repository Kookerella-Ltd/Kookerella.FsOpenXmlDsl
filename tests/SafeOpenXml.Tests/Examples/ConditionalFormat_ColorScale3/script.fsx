#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number 10.0) ]; row [ cell (Number 50.0) ]; row [ cell (Number 90.0) ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A3"), ColorScale3(Rgb(255uy, 0uy, 0uy), Rgb(255uy, 255uy, 0uy), Rgb(0uy, 128uy, 0uy))) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
