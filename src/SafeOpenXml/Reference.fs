namespace SafeOpenXml

open System
open System.Text.RegularExpressions

/// A single cell address within a worksheet grid, zero-based in both axes.
/// Row 0 / Col 0 is spreadsheet cell "A1".
[<Struct>]
type CellRef =
    { Row: int
      Col: int }

/// Conversions between `CellRef` and the "A1"-style addresses OOXML uses on the wire.
module CellRef =

    let create row col =
        if row < 0 then invalidArg (nameof row) "Row must be >= 0"
        if col < 0 then invalidArg (nameof col) "Col must be >= 0"
        { Row = row; Col = col }

    /// 0 -> "A", 25 -> "Z", 26 -> "AA", ...
    let columnLetters (col: int) : string =
        if col < 0 then invalidArg (nameof col) "Col must be >= 0"
        let rec loop (n: int) (acc: string) =
            if n < 0 then
                acc
            else
                let letter = char (int 'A' + n % 26)
                loop (n / 26 - 1) (string letter + acc)
        loop col ""

    /// "A" -> 0, "Z" -> 25, "AA" -> 26, ...
    let columnIndex (letters: string) : int =
        letters.ToUpperInvariant()
        |> Seq.fold (fun acc c -> acc * 26 + (int c - int 'A' + 1)) 0
        |> fun n -> n - 1

    /// Renders a `CellRef` as an OOXML cell reference, e.g. "B3".
    let toA1 (ref: CellRef) : string =
        sprintf "%s%d" (columnLetters ref.Col) (ref.Row + 1)

    let private a1Pattern = Regex(@"^\$?([A-Za-z]+)\$?(\d+)$", RegexOptions.Compiled)

    /// Parses an OOXML cell reference such as "B3" or "$B$3" back into a `CellRef`.
    let ofA1 (a1: string) : CellRef =
        let m = a1Pattern.Match(a1)
        if not m.Success then
            invalidArg (nameof a1) (sprintf "'%s' is not a valid cell reference" a1)
        { Row = int m.Groups.[2].Value - 1
          Col = columnIndex m.Groups.[1].Value }

    /// True row/column ordering (top-to-bottom, left-to-right), used to sort cells/rows for OOXML output.
    let compare (a: CellRef) (b: CellRef) =
        if a.Row <> b.Row then Operators.compare a.Row b.Row
        else Operators.compare a.Col b.Col
