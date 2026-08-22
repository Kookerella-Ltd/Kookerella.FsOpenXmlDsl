#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Widgets"); cell (Number 3.0); cell (Number 8.0); cell (Number 5.0); cell (Number 9.0) ]; row [ cell (Text "Gadgets"); cell (Number 6.0); cell (Number 4.0); cell (Number 7.0); cell (Number 2.0) ]; SparklineGroup { Style = { SparklineStyle.Default with ShowHigh = true; ShowLow = true }; Sparklines = [ { Cell = (CellRef.ofA1 "F1"); DataTopLeft = (CellRef.ofA1 "B1"); DataBottomRight = (CellRef.ofA1 "E1") }; { Cell = (CellRef.ofA1 "F2"); DataTopLeft = (CellRef.ofA1 "B2"); DataBottomRight = (CellRef.ofA1 "E2") } ] } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
