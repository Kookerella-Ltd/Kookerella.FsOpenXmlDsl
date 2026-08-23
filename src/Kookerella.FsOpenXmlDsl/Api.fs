namespace Kookerella.FsOpenXmlDsl

open System.IO
open Kookerella.FsOpenXmlDsl.Interpreter

/// Friendly entry points: render a `Workbook` to an .xlsx file/stream (the interpreter),
/// and parse an .xlsx file/stream back into a `Workbook` (the reverse transform).
module Workbook =

    let save (path: string) (wb: Workbook) : unit = Writer.saveToFile wb path

    let saveToStream (stream: Stream) (wb: Workbook) : unit = Writer.saveToStream wb stream

    let load (path: string) : Workbook = Reader.loadFromFile path

    let loadFromStream (stream: Stream) : Workbook = Reader.loadFromStream stream

    /// Renders `wb` as a self-contained F# script that, when run, rebuilds an equivalent
    /// file at `outputFileName`. `referenceLines` are whatever raw `#r` directives the
    /// caller needs so the script can locate the Kookerella.FsOpenXmlDsl assembly - this has no
    /// opinion on that, since it depends on where the script ends up living.
    let generateScript (referenceLines: string list) (outputFileName: string) (wb: Workbook) : string =
        CodeGen.generate referenceLines outputFileName wb
