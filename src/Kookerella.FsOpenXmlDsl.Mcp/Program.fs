module Kookerella.FsOpenXmlDsl.Mcp.Program

open System
open System.IO
open Microsoft.Extensions.DependencyInjection
open Microsoft.Extensions.Hosting
open Microsoft.Extensions.Logging

let private convertUsage =
    "Usage: fsopenxmldsl-mcp convert <input.xlsx> --lang <fsharp|csharp|xml|json> [-o <output-file>] [--rebuild-as <name.xlsx>]\n\n\
     Reads an existing Excel workbook and prints (or saves) an equivalent representation - \n\
     the same translation the generate_fsharp_script/generate_csharp_script/generate_xml/ \n\
     generate_json MCP tools expose to an agent, available here as a plain CLI for anyone \n\
     who isn't going through an MCP client.\n\n\
     Options:\n\
     \x20 --lang, -l          Required. \"fsharp\", \"csharp\", \"xml\", or \"json\".\n\
     \x20 -o, --output        Write the result to this file instead of stdout.\n\
     \x20 --rebuild-as        fsharp/csharp only - the filename the generated script itself\n\
     \x20                     saves its rebuilt workbook to. Defaults to \"output.xlsx\".\n\
     \x20                     Ignored for --lang xml/json, which have no script to embed a\n\
     \x20                     save path into - it's just the workbook's own data as XML/JSON."

let private buildUsage =
    "Usage: fsopenxmldsl-mcp build <input.xml|input.json> <output.xlsx>\n\n\
     Builds an Excel workbook from XML matching Xml.xsd, or JSON matching Json.schema.json \n\
     (see generate_xml/create_workbook_from_xml and generate_json/create_workbook_from_json) \n\
     and saves it to disk - the inverse of `convert --lang xml` / `convert --lang json`. The \n\
     natural target for a caller that already produces data as XML or JSON (e.g. an XSLT \n\
     pipeline, or a plain JSON-emitting script) and wants to reach Excel without writing any \n\
     code. Which format to read is inferred from <input>'s own extension (.xml or .json)."

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

    let validLangs = [ "fsharp"; "csharp"; "xml"; "json" ]

    match loop args (None, None, None, "output.xlsx") with
    | Error e -> Error e
    | Ok(None, _, _, _) -> Error "Missing <input.xlsx> path."
    | Ok(_, None, _, _) -> Error "Missing required --lang <fsharp|csharp|xml|json>."
    | Ok(Some path, Some lang, output, rebuildAs) when validLangs |> List.contains (lang.ToLowerInvariant()) ->
        Ok(path, lang.ToLowerInvariant(), output, rebuildAs)
    | Ok(_, Some lang, _, _) -> Error(sprintf "Unknown --lang value '%s' (expected fsharp, csharp, xml, or json)." lang)

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
                        match lang with
                        | "fsharp" -> WorkbookTools.GenerateFSharpScript(inputPath, rebuildAs)
                        | "csharp" -> WorkbookTools.GenerateCSharpScript(inputPath, rebuildAs)
                        | "json" -> WorkbookTools.GenerateJson(inputPath)
                        | _ -> WorkbookTools.GenerateXml(inputPath)

                    match output with
                    | Some path ->
                        File.WriteAllText(path, source)
                        eprintfn "Wrote %s" path
                    | None -> printfn "%s" source

                    0
                with ex ->
                    eprintfn "Conversion failed: %s" ex.Message
                    1

let private runBuild (args: string list) : int =
    match args with
    | ("-h" | "--help") :: _ ->
        printfn "%s" buildUsage
        0
    | [ inputPath; outputPath ] ->
        if not (File.Exists inputPath) then
            eprintfn "File not found: %s" inputPath
            1
        else
            match Path.GetExtension(inputPath).ToLowerInvariant() with
            | ".xml"
            | ".json" as ext ->
                try
                    let content = File.ReadAllText inputPath

                    let message =
                        if ext = ".json" then
                            WorkbookTools.CreateWorkbookFromJson(content, outputPath)
                        else
                            WorkbookTools.CreateWorkbookFromXml(content, outputPath)

                    eprintfn "%s" message
                    0
                with ex ->
                    eprintfn "Build failed: %s" ex.Message
                    1
            | other ->
                eprintfn "Unrecognized input extension '%s' (expected .xml or .json)." other
                1
    | _ ->
        eprintfn "%s" buildUsage
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
    | "build" :: rest -> runBuild rest
    | _ -> runServer argv
