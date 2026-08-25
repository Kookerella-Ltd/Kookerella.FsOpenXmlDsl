module Kookerella.FsOpenXmlDsl.Mcp.Program

open System
open System.IO
open Microsoft.Extensions.DependencyInjection
open Microsoft.Extensions.Hosting
open Microsoft.Extensions.Logging

let private convertUsage =
    "Usage: fsopenxmldsl-mcp convert <input.xlsx> --lang <fsharp|csharp> [-o <output-file>] [--rebuild-as <name.xlsx>]\n\n\
     Reads an existing Excel workbook and prints (or saves) equivalent, runnable source that \n\
     rebuilds an equivalent file - the same translation the generate_fsharp_script/generate_csharp_script \n\
     MCP tools expose to an agent, available here as a plain CLI for anyone who isn't going through an \n\
     MCP client.\n\n\
     Options:\n\
     \x20 --lang, -l          Required. \"fsharp\" or \"csharp\".\n\
     \x20 -o, --output        Write the generated source to this file instead of stdout.\n\
     \x20 --rebuild-as        The filename the generated script itself saves its rebuilt workbook to.\n\
     \x20                     Defaults to \"output.xlsx\"."

/// Parses `convert`'s own arguments (everything after the `convert` token), separately from
/// argv as a whole - unlike an MCP tool call's already-typed JSON arguments, a CLI has to
/// parse a flat string array by hand.
let private parseConvertArgs (args: string list) =
    let rec loop remaining (inputPath, lang, output, rebuildAs) =
        match remaining with
        | [] -> Ok(inputPath, lang, output, rebuildAs)
        | ("--lang" | "-l") :: value :: rest -> loop rest (inputPath, Some value, output, rebuildAs)
        | ("-o" | "--output") :: value :: rest -> loop rest (inputPath, lang, Some value, rebuildAs)
        | "--rebuild-as" :: value :: rest -> loop rest (inputPath, lang, output, value)
        | flag :: _ when flag.StartsWith("-") -> Error(sprintf "Unrecognized option or missing value: %s" flag)
        | path :: rest when inputPath = None -> loop rest (Some path, lang, output, rebuildAs)
        | extra :: _ -> Error(sprintf "Unexpected argument: %s" extra)

    match loop args (None, None, None, "output.xlsx") with
    | Error e -> Error e
    | Ok(None, _, _, _) -> Error "Missing <input.xlsx> path."
    | Ok(_, None, _, _) -> Error "Missing required --lang <fsharp|csharp>."
    | Ok(Some path, Some lang, output, rebuildAs) when
        lang.Equals("fsharp", StringComparison.OrdinalIgnoreCase)
        || lang.Equals("csharp", StringComparison.OrdinalIgnoreCase)
        ->
        Ok(path, lang.ToLowerInvariant(), output, rebuildAs)
    | Ok(_, Some lang, _, _) -> Error(sprintf "Unknown --lang value '%s' (expected fsharp or csharp)." lang)

let private runConvert (args: string list) : int =
    match args with
    | ("-h" | "--help") :: _ ->
        printfn "%s" convertUsage
        0
    | _ ->
        match parseConvertArgs args with
        | Error message ->
            eprintfn "%s\n\n%s" message convertUsage
            1
        | Ok(inputPath, lang, output, rebuildAs) ->
            if not (File.Exists inputPath) then
                eprintfn "File not found: %s" inputPath
                1
            else
                try
                    let source =
                        if lang = "fsharp" then
                            WorkbookTools.GenerateFSharpScript(inputPath, rebuildAs)
                        else
                            WorkbookTools.GenerateCSharpScript(inputPath, rebuildAs)

                    match output with
                    | Some path ->
                        File.WriteAllText(path, source)
                        eprintfn "Wrote %s" path
                    | None -> printfn "%s" source

                    0
                with ex ->
                    eprintfn "Conversion failed: %s" ex.Message
                    1

let private runServer (argv: string[]) : int =
    let builder = Host.CreateApplicationBuilder(argv)

    // MCP over stdio reserves stdout entirely for the JSON-RPC protocol stream - any log
    // output has to go to stderr instead, or it corrupts the protocol from the client's
    // point of view.
    builder.Logging.AddConsole(fun options -> options.LogToStandardErrorThreshold <- LogLevel.Trace)
    |> ignore

    builder.Services.AddMcpServer().WithStdioServerTransport().WithToolsFromAssembly()
    |> ignore

    builder.Build().RunAsync().GetAwaiter().GetResult()
    0

[<EntryPoint>]
let main argv =
    match Array.toList argv with
    | "convert" :: rest -> runConvert rest
    | _ -> runServer argv
