#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number -2.0); cell (Number 4.0); cell (Number -1.0); cell (Number 3.0) ]; SparklineGroup { Style = { SparklineStyle.Default with Type = Column; Color = (Some (Rgb(0uy, 112uy, 192uy))); ShowNegative = true }; Sparklines = [ { Cell = (CellRef.ofA1 "E1"); DataTopLeft = (CellRef.ofA1 "A1"); DataBottomRight = (CellRef.ofA1 "D1") } ] } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
