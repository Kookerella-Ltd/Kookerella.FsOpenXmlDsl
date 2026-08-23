open System
open Kookerella.FsOpenXmlDsl
open type Kookerella.FsOpenXmlDsl.SheetDsl

let headerStyle: CellStyle =
    { CellStyle.Default with
        Font = Some { FontStyle.Default with Bold = true; Color = Some Color.white }
        Fill = Some { Color = Rgb(68uy, 84uy, 106uy) } }

let currency = { CellStyle.Default with NumberFormat = Some Currency }

let buildInvoice () : Workbook =
    let invoice =
        sheet
            "Invoice"
            [ row [ cell (Text "Item", style = headerStyle)
                    cell (Text "Qty", style = headerStyle)
                    cell (Text "Unit Price", style = headerStyle)
                    cell (Text "Total", style = headerStyle) ]
              row [ cell (Text "Widgets")
                    cell (Number 4.0)
                    cell (Number 2.5, style = currency)
                    cell (Formula("B2*C2", Some 10.0), style = currency) ]
              row [ cell (Text "Gadgets")
                    cell (Number 2.0)
                    cell (Number 19.99, style = currency)
                    cell (Formula("B3*C3", Some 39.98), style = currency) ]
              // index = 4 skips row 4 (index 3) entirely, and col = 2 skips column B -
              // landing the date at C5 instead of the next sequential row/column.
              row ([ cell (Text "Invoice date"); cell (Date DateTime.Today, col = 2) ], index = 4)
              ColumnWidth(0, 18.0)
              ColumnWidth(3, 14.0)
              Freeze(1, 0) ]

    workbook [ invoice ]

[<EntryPoint>]
let main argv =
    let path =
        match argv with
        | [| p |] -> p
        | _ -> IO.Path.Combine(IO.Path.GetTempPath(), "fsopenxmldsl-sample.xlsx")

    let wb = buildInvoice ()
    Workbook.save path wb
    printfn "Wrote %s" path

    // Reverse transform: read the file we just wrote back into the DSL.
    let roundTripped = Workbook.load path
    let sheet = roundTripped.Sheets.[0]
    printfn "Read back sheet '%s' with %d cells:" sheet.Name sheet.Cells.Length

    sheet.Cells
    |> List.sortBy (fun c -> c.Ref.Row, c.Ref.Col)
    |> List.iter (fun c -> printfn "  %s = %A" (CellRef.toA1 c.Ref) c.Value)

    0
