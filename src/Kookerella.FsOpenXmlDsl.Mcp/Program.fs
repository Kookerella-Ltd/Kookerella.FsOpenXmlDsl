module Kookerella.FsOpenXmlDsl.Mcp.Program

open Microsoft.Extensions.DependencyInjection
open Microsoft.Extensions.Hosting
open Microsoft.Extensions.Logging

[<EntryPoint>]
let main argv =
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
