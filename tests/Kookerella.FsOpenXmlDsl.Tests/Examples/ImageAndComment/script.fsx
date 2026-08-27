#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Logo below:") ]; comment ((CellRef.ofA1 "A1"), "See the logo below for the current brand mark."); EmbeddedImage { Data = System.Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7"); Format = Gif; TopLeftAnchor = (CellRef.ofA1 "A3"); BottomRightAnchor = (CellRef.ofA1 "C10") } ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
