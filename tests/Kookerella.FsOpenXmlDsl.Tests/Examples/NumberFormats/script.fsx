#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Integer"); cell (Number 1234.0, style = { CellStyle.Default with NumberFormat = (Some (Integer)) }) ]; row [ cell (Text "TwoDecimal"); cell (Number 1234.5, style = { CellStyle.Default with NumberFormat = (Some (TwoDecimal)) }) ]; row [ cell (Text "Percentage"); cell (Number 0.42, style = { CellStyle.Default with NumberFormat = (Some (Percentage)) }) ]; row [ cell (Text "Currency"); cell (Number 19.99, style = { CellStyle.Default with NumberFormat = (Some (Currency)) }) ]; row [ cell (Text "ShortDate"); cell (Date (System.DateTime.FromOADate(46082.0)), style = { CellStyle.Default with NumberFormat = (Some (ShortDate)) }) ]; row [ cell (Text "DateAndTime"); cell (Date (System.DateTime.FromOADate(46082.5625)), style = { CellStyle.Default with NumberFormat = (Some (DateAndTime)) }) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
