#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Apple") ]; row [ cell (Text "Banana") ]; row [ cell (Text "Apple") ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A3"), DuplicateValuesRule({ CellStyle.Default with Fill = (Some ({ Color = Rgb(255uy, 199uy, 206uy) })) })) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
