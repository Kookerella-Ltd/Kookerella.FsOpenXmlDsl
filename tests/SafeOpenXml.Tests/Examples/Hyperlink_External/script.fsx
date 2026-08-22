#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\SafeOpenXml.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\SafeOpenXml.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open SafeOpenXml
open type SafeOpenXml.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Open-XML-SDK on GitHub") ]; hyperlink ((CellRef.ofA1 "A1"), ExternalHyperlink "https://github.com/dotnet/Open-XML-SDK", tooltip = "Open in browser", display = "dotnet/Open-XML-SDK") ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
