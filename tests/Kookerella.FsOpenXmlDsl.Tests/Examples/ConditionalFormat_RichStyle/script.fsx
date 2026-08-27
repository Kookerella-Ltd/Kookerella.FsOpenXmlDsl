#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Number 5.0) ]; conditionalFormat ((CellRef.ofA1 "A1"), (CellRef.ofA1 "A1"), CellValueRule(GreaterThan, "0", None, { CellStyle.Default with Font = (Some ({ FontStyle.Default with Bold = true })); Fill = (Some ({ Color = Rgb(255uy, 0uy, 0uy) })); Border = (Some ({ BorderStyle.None with Bottom = (Some ({ Style = Thin; Color = (Some (Rgb(0uy, 0uy, 0uy))) })) })); Alignment = (Some ({ AlignmentStyle.Default with Horizontal = (Some (AlignCenter)) })) })) ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
