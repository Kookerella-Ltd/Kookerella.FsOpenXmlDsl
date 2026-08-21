namespace SafeOpenXml

open System.IO
open SafeOpenXml.Interpreter

/// Friendly entry points: render a `Workbook` to an .xlsx file/stream (the interpreter),
/// and parse an .xlsx file/stream back into a `Workbook` (the reverse transform).
module Workbook =

    let save (path: string) (wb: Workbook) : unit = Writer.saveToFile wb path

    let saveToStream (stream: Stream) (wb: Workbook) : unit = Writer.saveToStream wb stream

    let load (path: string) : Workbook = Reader.loadFromFile path

    let loadFromStream (stream: Stream) : Workbook = Reader.loadFromStream stream
