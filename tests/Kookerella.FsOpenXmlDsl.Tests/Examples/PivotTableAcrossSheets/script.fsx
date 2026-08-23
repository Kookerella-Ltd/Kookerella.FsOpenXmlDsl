#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Data" [ row [ cell (Text "Category"); cell (Text "Amount") ]; row [ cell (Text "A"); cell (Number 3.0) ]; row [ cell (Text "B"); cell (Number 7.0) ]; row [ cell (Text "A"); cell (Number 4.0) ] ]
let sheet1 = sheet "Report" [ row [ cell (Text "Pivot table below:") ]; EmbeddedPivotTable { SourceSheet = (Some ("Data")); SourceTopLeft = (CellRef.ofA1 "A1"); SourceBottomRight = (CellRef.ofA1 "B4"); RowField = "Category"; ColumnField = None; ValueField = "Amount"; Aggregation = PivotCount; ValueCaption = None; TopLeftAnchor = (CellRef.ofA1 "A3") } ]

let wb = workbook [ sheet0; sheet1 ]

wb |> Workbook.save "output.xlsx"
