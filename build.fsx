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

open System.Net.Http
open System.Text.Json
open System.Xml.Linq
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

// NuGet PackageId defaults to the project filename for all three - see each .fsproj/
// .csproj's own comment on why that's intentionally not overridden.
let fsCorePackageId = "Kookerella.FsOpenXmlDsl"
let csWrapperPackageId = "Kookerella.CsOpenXmlDsl"
let mcpPackageId = "Kookerella.FsOpenXmlDsl.Mcp"

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

let private httpClient = new HttpClient()

let private getStringSync (url: string) : string =
    httpClient.GetStringAsync(url) |> Async.AwaitTask |> Async.RunSynchronously

/// Every version ever published for a package, oldest first - queries nuget.org's own
/// flatcontainer index directly (the same one CLAUDE.md already points at for "has this
/// version finished indexing yet") rather than assuming anything about local state, since
/// the whole point of this check is to catch drift between what's published and what's in
/// the repo.
let private publishedVersions (packageId: string) : string list =
    let url = sprintf "https://api.nuget.org/v3-flatcontainer/%s/index.json" (packageId.ToLowerInvariant())
    use doc = JsonDocument.Parse(getStringSync url)
    doc.RootElement.GetProperty("versions").EnumerateArray() |> Seq.map (fun v -> v.GetString()) |> List.ofSeq

let private latestPublishedVersion (packageId: string) : string = publishedVersions packageId |> List.last

/// The exact `<nupkg-id>.<version>.nupkg` filename's version segment - reading it back out
/// of the packed file rather than re-parsing the `.fsproj`/`.csproj` keeps this in sync with
/// whatever `Pack` actually produced, not whatever the source file says (the two are only
/// guaranteed to agree if `Pack` already ran).
let private versionOfNupkg (packageId: string) (nupkgPath: string) : string =
    let fileName = System.IO.Path.GetFileNameWithoutExtension nupkgPath
    fileName.Substring(packageId.Length + 1)

/// Fetches one published version's raw `.nuspec` XML directly from the flatcontainer -
/// this is "what NuGet actually served a consumer," not what the repo's own `.fsproj` says,
/// which is exactly the distinction the font-ordering-adjacent wrapper-staleness bug hinged
/// on: local source is always self-consistent, only *published* packages can drift.
let private fetchNuspec (packageId: string) (version: string) : string =
    let idLower = packageId.ToLowerInvariant()
    let url = sprintf "https://api.nuget.org/v3-flatcontainer/%s/%s/%s.nuspec" idLower version idLower
    getStringSync url

/// Every distinct minimum version a published nuspec declares for one dependency id, across
/// every `<group targetFramework="...">` (a `ProjectReference`-converted dependency is
/// usually duplicated once per TFM group with the same version, but this checks all of them
/// rather than assume the first one found is representative).
let private dependencyFloors (nuspecXml: string) (dependencyId: string) : string list =
    let doc = XDocument.Parse(nuspecXml)
    let ns = doc.Root.Name.Namespace

    doc.Descendants(ns + "dependency")
    |> Seq.filter (fun d -> (d.Attribute(XName.Get "id").Value).Equals(dependencyId, System.StringComparison.OrdinalIgnoreCase))
    |> Seq.map (fun d -> d.Attribute(XName.Get "version").Value)
    |> Seq.distinct
    |> List.ofSeq

/// Pushes one project's packed .nupkg to nuget.org - but first checks whether `packageId`
/// already has this exact version published, and skips (not fails) if so. This is what
/// makes `PublishAll` safe to run on *every* release regardless of which package(s) actually
/// changed: pushing a version NuGet already has would otherwise error out, which is exactly
/// why this repo fell into calling `PushCore`/`PushWrapper`/`PushMcp` individually - and
/// individually is how the wrapper's dependency on the core silently went stale for three
/// releases in a row. Always use `PublishAll`, never one of the three `Push*` targets alone
/// (see `VerifyDependencyFreshness` below for the other half of this fix).
///
/// Deliberately bypasses `DotNet.exec`/FAKE's own process tracing entirely and shells out
/// via a raw `Process` instead for the actual push - FAKE logs a command's full argument
/// list to the console on every invocation (success or failure, see the `.> "dotnet.exe"
/// ...` lines other targets print), which would put the API key in plain text in the build
/// output. Reads the key from NUGET_API_KEY rather than a parameter so it never appears in
/// FAKE's own target-invocation logging either. `dotnet nuget push`'s own stdout/stderr
/// never echoes the key back, so relaying those verbatim is safe.
let push (packageId: string) (projPath: string) =
    let nupkg = findNupkg projPath
    let version = versionOfNupkg packageId nupkg

    if publishedVersions packageId |> List.contains version then
        Trace.tracefn "%s %s is already published - skipping (this package didn't change this release)." packageId version
    else

    let apiKey =
        match Environment.environVarOrNone "NUGET_API_KEY" with
        | Some key -> key
        | None -> failwith "NUGET_API_KEY is not set - export it before running a Push target."

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

Target.create "PushCore" (fun _ -> push fsCorePackageId fsCoreProj)
Target.create "PushWrapper" (fun _ -> push csWrapperPackageId csWrapperProj)
Target.create "PushMcp" (fun _ -> push mcpPackageId mcpProj)

// The other half of the wrapper-staleness fix, alongside `push`'s own skip-if-published
// idempotency: independent of *this* release having remembered to run `PublishAll`
// correctly, directly check the live, already-published state on nuget.org and fail loudly
// if it's still wrong - e.g. because someone ran `PushCore` alone again out of habit, the
// exact way this drifted the first time. Only checks the wrapper's dependency on the core -
// the Mcp tool bundles its own dependencies as a self-contained `dotnet tool` payload rather
// than declaring them in its nuspec at all (verified: its published nuspec has no
// `<dependencies>` section), so there's no "stale floor" for it to have.
Target.create "VerifyDependencyFreshness" (fun _ ->
    let latestCore = latestPublishedVersion fsCorePackageId
    let latestWrapper = latestPublishedVersion csWrapperPackageId
    let wrapperNuspec = fetchNuspec csWrapperPackageId latestWrapper

    match dependencyFloors wrapperNuspec fsCorePackageId with
    | [] ->
        failwithf
            "%s %s has no declared dependency on %s at all - has the ProjectReference been removed?"
            csWrapperPackageId
            latestWrapper
            fsCorePackageId
    | floors when floors |> List.contains latestCore |> not ->
        failwithf
            "%s %s declares %s %s, but the latest published %s is %s. Bump and republish \
             %s (via PublishAll, even with no code changes) to refresh this - see the \
             'push' function's own doc comment in build.fsx for why this class of drift \
             happens at all."
            csWrapperPackageId
            latestWrapper
            fsCorePackageId
            (String.concat " / " floors)
            fsCorePackageId
            latestCore
            csWrapperPackageId
    | _ -> Trace.tracefn "%s %s correctly depends on the latest published %s %s." csWrapperPackageId latestWrapper fsCorePackageId latestCore)

Target.create "PublishAll" ignore

"Clean" ==> "Restore" ==> "Build" ==> "TestFast" ==> "TestSlow" |> ignore

// Every Pack target depends on the full test gate, not just Build - so `fake build -t
// PushMcp` on its own still refuses to run against a failing suite, the same as the
// combined PublishAll target does.
"TestSlow" ==> "PackCore" ==> "PushCore" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackWrapper" ==> "PushWrapper" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackMcp" ==> "PushMcp" ==> "PublishAll" |> ignore

// PublishAll's whole purpose is to be the *only* sanctioned way to push a release (push
// itself is written to be a no-op for whichever package(s) didn't change this time,
// specifically so there's never a reason to reach for PushCore/PushWrapper/PushMcp alone) -
// this final check confirms that policy actually held, using nuget.org's own live state
// rather than trusting that it did.
"PublishAll" ==> "VerifyDependencyFreshness" |> ignore

Target.runOrDefaultWithArguments "Build"
