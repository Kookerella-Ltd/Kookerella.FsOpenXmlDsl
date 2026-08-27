#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\Kookerella.FsOpenXmlDsl.dll"
#r "C:\\Users\\m_r_n\\source\\repos\\SafeOpenXml\\tests\\Kookerella.FsOpenXmlDsl.Tests\\bin\\Debug\\net10.0\\DocumentFormat.OpenXml.dll"

open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let sheet0 = sheet "Sheet1" [ row [ cell (Text "Centered and unlocked", style = { CellStyle.Default with Alignment = (Some ({ AlignmentStyle.Default with Horizontal = (Some (AlignCenter)); Vertical = (Some (AlignMiddle)); WrapText = true })); Protection = (Some ({ CellProtection.Default with Locked = false; Hidden = true })) }) ] ]

let wb = workbook [ sheet0 ]

wb |> Workbook.save "output.xlsx"
