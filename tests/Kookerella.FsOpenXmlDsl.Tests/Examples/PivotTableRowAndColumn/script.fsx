#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Region"); cell (Text "Quarter"); cell (Text "Sales") ]; row [ cell (Text "East"); cell (Text "Q1"); cell (Number 10.0) ]; row [ cell (Text "East"); cell (Text "Q2"); cell (Number 5.0) ]; row [ cell (Text "West"); cell (Text "Q1"); cell (Number 20.0) ]; row [ cell (Text "West"); cell (Text "Q2"); cell (Number 15.0) ]; EmbeddedPivotTable { SourceSheet = None; SourceTopLeft = (CellRef.ofA1 "A1"); SourceBottomRight = (CellRef.ofA1 "C5"); RowField = "Region"; ColumnField = (Some ("Quarter")); ValueField = "Sales"; Aggregation = PivotSum; ValueCaption = (Some ("Total Sales")); TopLeftAnchor = (CellRef.ofA1 "E1") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
