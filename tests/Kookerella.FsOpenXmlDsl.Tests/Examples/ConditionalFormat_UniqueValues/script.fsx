#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Apple") ]; row [ cell (Text "Banana") ]; row [ cell (Text "Apple") ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A3"), UniqueValuesRule({ CellStyle.Default with Fill = (Some ({ Color = Rgb(198uy, 239uy, 206uy) })) })) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
