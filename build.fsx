// The fake-cli runner (as of 6.1.4, the latest at time of writing) still runs on
// FSharp.Core 8.0.0 - without pinning it explicitly here, Paket resolves the Fake packages'
// transitive dependency on a newer FSharp.Core the runner can't load. See
// https://github.com/fsharp/FAKE/issues/2001.
#r "paket:
nuget Fake.Core.Target 6.1.3
nuget Fake.Core.Environment 6.1.3
nuget Fake.DotNet.Cli 6.1.3
nuget Fake.IO.FileSystem 6.1.3
nuget FSharp.Core 8.0.401 //"

open Fake.Core
open Fake.DotNet
open Fake.IO
open Fake.IO.FileSystemOperators
open Fake.IO.Globbing.Operators
open Fake.Core.TargetOperators

// This codifies the exact release sequence that was, until now, run by hand across a chat
// session: build, run both test suites, pack the three packages, and push whichever ones
// changed. It deliberately does NOT bump version numbers or touch the MCP Registry -
// versions are still edited by hand in each .csproj/.fsproj first (a semver judgment call,
// not something to automate), and `mcp-publisher publish` needs an interactive GitHub
// device-flow login (`mcp-publisher login github`) that can't run inside a script; run that
// step by hand afterward, same as before.

let solution = "Kookerella.FsOpenXmlDsl.slnx"

let fsCoreProj = "src/Kookerella.FsOpenXmlDsl/Kookerella.FsOpenXmlDsl.fsproj"
let csWrapperProj = "src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj"
let mcpProj = "src/Kookerella.FsOpenXmlDsl.Mcp/Kookerella.FsOpenXmlDsl.Mcp.fsproj"

let fsTestsProj = "tests/Kookerella.FsOpenXmlDsl.Tests/Kookerella.FsOpenXmlDsl.Tests.fsproj"
let csTestsProj = "tests/Kookerella.CsOpenXmlDsl.Tests/Kookerella.CsOpenXmlDsl.Tests.csproj"

let releaseDir (projPath: string) = (Fake.IO.Path.getDirectory projPath) @@ "bin" @@ "Release"

/// The single .nupkg `dotnet pack` produces for a project - fails loudly (rather than
/// silently picking one) if Pack hasn't been run or somehow produced more than one, since
/// either would mean Push is about to do the wrong thing.
let findNupkg (projPath: string) : string =
    match !!(releaseDir projPath @@ "*.nupkg") |> List.ofSeq with
    | [ single ] -> single
    | [] -> failwithf "No .nupkg found under %s - run the Pack target first." (releaseDir projPath)
    | many -> failwithf "Expected exactly one .nupkg under %s, found %d: %A" (releaseDir projPath) many.Length many

let dotnet (args: string) (workingDir: string) =
    let result =
        DotNet.exec (fun opts -> { opts with WorkingDirectory = workingDir }) "" args

    if not result.OK then
        failwithf "'dotnet %s' failed in %s (exit %d)" args workingDir result.ExitCode

let runTests (filter: string) (projPaths: string list) =
    for proj in projPaths do
        dotnet (sprintf "test \"%s\" --filter \"%s\"" proj filter) "."

let pack (projPath: string) =
    dotnet (sprintf "pack \"%s\" -c Release" projPath) "."

/// Pushes one project's packed .nupkg to nuget.org. Deliberately bypasses `DotNet.exec`/
/// FAKE's own process tracing entirely and shells out via a raw `Process` instead - FAKE
/// logs a command's full argument list to the console on every invocation (success or
/// failure, see the `.> "dotnet.exe" ...` lines other targets print), which would put the
/// API key in plain text in the build output. Reads the key from NUGET_API_KEY rather than
/// a parameter so it never appears in FAKE's own target-invocation logging either.
/// `dotnet nuget push`'s own stdout/stderr never echoes the key back, so relaying those
/// verbatim is safe.
let push (projPath: string) =
    let apiKey =
        match Environment.environVarOrNone "NUGET_API_KEY" with
        | Some key -> key
        | None -> failwith "NUGET_API_KEY is not set - export it before running a Push target."

    let nupkg = findNupkg projPath
    let source = "https://api.nuget.org/v3/index.json"

    Trace.tracefn "Pushing %s to %s (api-key redacted)..." nupkg source

    let psi = System.Diagnostics.ProcessStartInfo("dotnet")
    psi.RedirectStandardOutput <- true
    psi.RedirectStandardError <- true
    psi.UseShellExecute <- false

    for a in [ "nuget"; "push"; nupkg; "--api-key"; apiKey; "--source"; source ] do
        psi.ArgumentList.Add(a)

    use proc = System.Diagnostics.Process.Start(psi)
    let stdout = proc.StandardOutput.ReadToEnd()
    let stderr = proc.StandardError.ReadToEnd()
    proc.WaitForExit()

    if stdout <> "" then Trace.log stdout
    if stderr <> "" then Trace.log stderr

    if proc.ExitCode <> 0 then
        failwithf "dotnet nuget push failed for %s (exit %d) - see output above." nupkg proc.ExitCode

Target.create "Clean" (fun _ ->
    !! "src/**/bin" ++ "src/**/obj" ++ "tests/**/bin" ++ "tests/**/obj" ++ "samples/**/bin" ++ "samples/**/obj"
    |> Shell.cleanDirs)

Target.create "Restore" (fun _ -> dotnet (sprintf "restore \"%s\"" solution) ".")

Target.create "Build" (fun _ -> dotnet (sprintf "build \"%s\" --no-restore" solution) ".")

Target.create "TestFast" (fun _ -> runTests "Category!=Slow" [ fsTestsProj; csTestsProj ])

// The slow group actually shells out to `dotnet run`/`dotnet fsi` per generated example
// script - see each test project's own comment on why that's its own category rather than
// part of the default run.
Target.create "TestSlow" (fun _ -> runTests "Category=Slow" [ fsTestsProj; csTestsProj ])

Target.create "PackCore" (fun _ -> pack fsCoreProj)
Target.create "PackWrapper" (fun _ -> pack csWrapperProj)
Target.create "PackMcp" (fun _ -> pack mcpProj)

Target.create "PushCore" (fun _ -> push fsCoreProj)
Target.create "PushWrapper" (fun _ -> push csWrapperProj)
Target.create "PushMcp" (fun _ -> push mcpProj)

Target.create "PublishAll" ignore

"Clean" ==> "Restore" ==> "Build" ==> "TestFast" ==> "TestSlow" |> ignore

// Every Pack target depends on the full test gate, not just Build - so `fake build -t
// PushMcp` on its own still refuses to run against a failing suite, the same as the
// combined PublishAll target does.
"TestSlow" ==> "PackCore" ==> "PushCore" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackWrapper" ==> "PushWrapper" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackMcp" ==> "PushMcp" ==> "PublishAll" |> ignore

Target.runOrDefaultWithArguments "Build"
