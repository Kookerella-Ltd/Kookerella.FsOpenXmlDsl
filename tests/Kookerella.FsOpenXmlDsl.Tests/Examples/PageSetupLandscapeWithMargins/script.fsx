#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Wide report") ]; PageSetup({ PageSetup.Default with Orientation = Landscape; PaperSize = (Some (A4)); Scaling = (Some (ScalePercent 85)); Margins = { PageMargins.Default with Left = 0.5; Right = 0.5 }; Header = (Some ("&C&\"Arial,Bold\"Quarterly Report")); Footer = (Some ("&LPage &P of &N&R&D")) }) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
