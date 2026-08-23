#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Region"); cell (Text "Sales") ]; row [ cell (Text "East"); cell (Number 10.0) ]; row [ cell (Text "West"); cell (Number 20.0) ]; row [ cell (Text "East"); cell (Number 5.0) ]; row [ cell (Text "West"); cell (Number 15.0) ]; EmbeddedPivotTable { SourceSheet = None; SourceTopLeft = (CellRef.ofA1 "A1"); SourceBottomRight = (CellRef.ofA1 "B5"); RowField = "Region"; ColumnField = None; ValueField = "Sales"; Aggregation = PivotSum; ValueCaption = None; TopLeftAnchor = (CellRef.ofA1 "D1") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
