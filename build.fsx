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
open System.Text.Json.Nodes
open System.Xml.Linq
open System.IO.Compression
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

let mcpServerJson = "src/Kookerella.FsOpenXmlDsl.Mcp/.mcp/server.json"

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

/// The <Version> a project file declares locally. Unlike everything else in this file that
/// reads a *published* version, this reads local source directly - the one thing it checks
/// (server.json's own version fields, below) only needs to match this repo's own .fsproj,
/// not anything on nuget.org yet.
let private localProjectVersion (projPath: string) : string =
    let doc = XDocument.Load(projPath: string)
    let ns = doc.Root.Name.Namespace
    doc.Descendants(ns + "Version") |> Seq.map (fun e -> e.Value) |> Seq.head

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

/// Downloads one published version's raw `.nupkg` (just a zip file) into a temp file and
/// returns the path - the caller is responsible for deleting it.
let private downloadNupkg (packageId: string) (version: string) : string =
    let idLower = packageId.ToLowerInvariant()
    let url = sprintf "https://api.nuget.org/v3-flatcontainer/%s/%s/%s.%s.nupkg" idLower version idLower version
    let bytes = httpClient.GetByteArrayAsync(url) |> Async.AwaitTask |> Async.RunSynchronously
    let path = System.IO.Path.GetTempFileName()
    System.IO.File.WriteAllBytes(path, bytes)
    path

/// Reads one bundled assembly's own version out of a `dotnet tool` package's `.nupkg`,
/// without ever loading or executing it - `PackAsTool` bundles its full dependency closure
/// as plain files under `tools/<tfm>/any/`, rather than declaring them in the nuspec the way
/// a normal library package does (verified: `Kookerella.FsOpenXmlDsl.Mcp`'s own published
/// nuspec has no `<dependencies>` section at all), so this is the only way to see what a
/// published tool actually contains. `AssemblyName.GetAssemblyName` only reads metadata from
/// the file - the same "never load a foreign assembly to inspect it" caution this repo
/// already applies elsewhere to untrusted input.
let private bundledAssemblyVersion (nupkgPath: string) (dllFileName: string) : System.Version option =
    use archive = ZipFile.OpenRead(nupkgPath)

    archive.Entries
    |> Seq.tryFind (fun e -> e.Name.Equals(dllFileName, System.StringComparison.OrdinalIgnoreCase))
    |> Option.map (fun entry ->
        let tempDll = System.IO.Path.GetTempFileName() + ".dll"

        try
            entry.ExtractToFile(tempDll, true)
            System.Reflection.AssemblyName.GetAssemblyName(tempDll).Version
        finally
            System.IO.File.Delete tempDll)

/// MSBuild pads a NuGet package's 3-part `<Version>` into a 4-part `AssemblyVersion` with a
/// trailing `.0` - compare only the first three parts, not the padding.
let private assemblyVersionMatchesPackageVersion (assemblyVersion: System.Version) (packageVersion: string) : bool =
    match System.Version.TryParse(packageVersion) with
    | false, _ -> false
    | true, parsed -> assemblyVersion.Major = parsed.Major && assemblyVersion.Minor = parsed.Minor && assemblyVersion.Build = parsed.Build

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
// exact way this drifted the first time.
let private verifyWrapperDependencyFloor () =
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
    | _ -> Trace.tracefn "%s %s correctly depends on the latest published %s %s." csWrapperPackageId latestWrapper fsCorePackageId latestCore

// The Mcp tool's own version of the exact same problem, one layer deeper: it bundles its
// full dependency closure as plain files (`PackAsTool`, no nuspec `<dependencies>` at all -
// verified against a real published package) rather than declaring a NuGet floor, so it went
// stale the same way and for the same reason (built via `ProjectReference` against whatever
// local source existed the last time *it* was packed) - discovered live, while re-decompiling
// a demo template: the installed `fsopenxmldsl-mcp` 0.6.0 turned out to bundle
// Kookerella.FsOpenXmlDsl 0.3.0.0 (before the font-ordering fix) and Kookerella.CsOpenXmlDsl
// 0.3.1.0 (before its own floor refresh), months-old despite being "latest" at the time.
// `dotnet tool update` only reinstalls whatever's currently published - it can't fix a
// published package that's already stale, so this has to be caught (and the affected
// package republished) at release time instead.
let private verifyMcpBundleFreshness () =
    let latestCore = latestPublishedVersion fsCorePackageId
    let latestWrapper = latestPublishedVersion csWrapperPackageId
    let latestMcp = latestPublishedVersion mcpPackageId
    let nupkg = downloadNupkg mcpPackageId latestMcp

    try
        let checkBundled (dllFileName: string) (bundledPackageId: string) (latestVersion: string) =
            match bundledAssemblyVersion nupkg dllFileName with
            | None ->
                failwithf "%s %s doesn't bundle %s at all - has the ProjectReference been removed?" mcpPackageId latestMcp dllFileName
            | Some assemblyVersion when not (assemblyVersionMatchesPackageVersion assemblyVersion latestVersion) ->
                failwithf
                    "%s %s bundles %s version %O, but the latest published %s is %s. Bump \
                     and republish %s (via PublishAll, even with no code changes) to refresh \
                     this - see the 'push' function's own doc comment in build.fsx for why \
                     this class of drift happens at all."
                    mcpPackageId
                    latestMcp
                    dllFileName
                    assemblyVersion
                    bundledPackageId
                    latestVersion
                    mcpPackageId
            | Some assemblyVersion ->
                Trace.tracefn "%s %s correctly bundles %s %O, matching the latest published %s." mcpPackageId latestMcp dllFileName assemblyVersion bundledPackageId

        checkBundled "Kookerella.FsOpenXmlDsl.dll" fsCorePackageId latestCore
        checkBundled "Kookerella.CsOpenXmlDsl.dll" csWrapperPackageId latestWrapper
    finally
        System.IO.File.Delete nupkg

// Flagged as an unchecked gap in CLAUDE.md's own "Keep these in sync" table: the MCP
// registry publish rejects a version it can't find on NuGet, but nothing previously
// confirmed server.json's own two version fields (top-level and packages[0].version) agree
// with each other or with the .fsproj they're supposed to mirror by hand. Purely a local
// check - no network call, no dependency on anything having been published yet.
let private verifyServerJsonVersion () =
    let projVersion = localProjectVersion mcpProj
    use doc = JsonDocument.Parse(System.IO.File.ReadAllText mcpServerJson)
    let root = doc.RootElement
    let topLevelVersion = root.GetProperty("version").GetString()
    let packageVersion = root.GetProperty("packages").EnumerateArray() |> Seq.head |> fun p -> p.GetProperty("version").GetString()

    if topLevelVersion <> projVersion || packageVersion <> projVersion then
        failwithf
            "%s's version fields (top-level %s, packages[0].version %s) don't both match \
             %s's <Version> (%s) - these three are edited by hand, not derived, so update \
             whichever one(s) are stale."
            mcpServerJson
            topLevelVersion
            packageVersion
            mcpProj
            projVersion
    else
        Trace.tracefn "%s's version fields correctly match %s's <Version> (%s)." mcpServerJson mcpProj projVersion

let verifyDependencyFreshness () =
    verifyWrapperDependencyFloor ()
    verifyMcpBundleFreshness ()
    verifyServerJsonVersion ()

// Independently invocable (`fake run build.fsx -t VerifyDependencyFreshness --single-target`)
// as a standalone health check, any time - the two dependency-freshness checks need nothing
// but nuget.org's own already-published state, and the server.json check needs nothing but
// local source, so none of PublishAll's other prerequisites are required to run this alone.
Target.create "VerifyDependencyFreshness" (fun _ -> verifyDependencyFreshness ())

// PublishAll's whole purpose is to be the *only* sanctioned way to push a release (push
// itself is written to be a no-op for whichever package(s) didn't change this time,
// specifically so there's never a reason to reach for PushCore/PushWrapper/PushMcp alone) -
// running the same freshness check here, as PublishAll's own action rather than merely a
// same-named target chained *after* it (`"PublishAll" ==> "VerifyDependencyFreshness"` would
// make VerifyDependencyFreshness depend on PublishAll, not the other way round - running
// `-t PublishAll` would never reach it), is what actually confirms the policy held on every
// real `-t PublishAll` invocation, using nuget.org's own live state rather than trusting
// that it did.
Target.create "PublishAll" (fun _ -> verifyDependencyFreshness ())

// A second, additional distribution channel for the Mcp server, alongside (not replacing)
// the `dotnet tool` package PushMcp/PublishAll publish: a self-contained, single-file
// build per platform that needs no .NET runtime pre-installed at all, for exactly the
// audience `dotnet tool install` can't serve. Native AOT (smaller, faster) was tried first
// and rejected - not a maybe, an empirically confirmed crash: F#'s own sprintf/printf/
// failwithf machinery parses format strings and builds typed handlers via reflection
// (`MakeGenericMethod`) at runtime, which Native AOT's trimmer can't resolve statically.
// The ILC compile step itself succeeded and produced a native executable; running it
// against a real file crashed immediately on the first sprintf-driven code path hit
// (CsCodeGen's own `#:package` line rendering). Since sprintf/failwithf are idiomatic and
// pervasive throughout ordinary F# code, not one isolated call site, fixing this for real
// AOT support would mean auditing and rewriting every such usage across the F# core - a
// large, separate undertaking, not a quick fix. Self-contained *without* AOT sidesteps the
// whole problem: it bundles the full JIT-capable runtime (no trimming involved at all), so
// sprintf's reflection works exactly as it does today - verified by actually running the
// published binary against a real file, not just checking it compiled.
let private selfContainedRids = [ "win-x64"; "win-arm64"; "linux-x64"; "linux-arm64"; "osx-x64"; "osx-arm64" ]

let private selfContainedDir = (Fake.IO.Path.getDirectory mcpProj) @@ "bin" @@ "Release" @@ "self-contained"

let private mcpExeName (rid: string) =
    if rid.StartsWith "win" then "Kookerella.FsOpenXmlDsl.Mcp.exe" else "Kookerella.FsOpenXmlDsl.Mcp"

/// PublishSingleFile bundles the whole self-contained output into one executable per RID
/// (~80-85MB, the full runtime included) rather than the ~250 loose files a plain
/// self-contained publish produces - confirmed this doesn't break anything (verified
/// against a real file) since nothing here does its own file-system introspection of the
/// app's own directory.
let private publishSelfContained (rid: string) =
    let outDir = selfContainedDir @@ rid
    dotnet (sprintf "publish \"%s\" -c Release -r %s --self-contained -p:PublishSingleFile=true -o \"%s\"" mcpProj rid outDir) "."

/// Zips just the one executable (renamed to the tool's own command name, not the assembly
/// name) - not the loose .pdb/.xml files publish also drops alongside it, which are debug
/// symbols/doc comments, not needed to run. Unix RIDs lose the executable bit across a zip -
/// checked empirically (parsed the zip's own central directory for the Unix mode bits) that
/// this isn't specific to .NET's ZipFile API: a Linux/macOS binary cross-compiled *on
/// Windows* has no Unix mode bits to begin with (NTFS doesn't track POSIX permissions at
/// all), so no zip tool run from this machine could preserve what was never there. Document
/// `chmod +x` as a required step for those platforms rather than reaching for a tar.gz
/// writer that wouldn't actually fix anything here.
let private archiveSelfContained (version: string) (rid: string) =
    let outDir = selfContainedDir @@ rid
    let exeName = mcpExeName rid
    let archivePath = selfContainedDir @@ sprintf "fsopenxmldsl-mcp-%s-%s.zip" version rid

    if System.IO.File.Exists archivePath then
        System.IO.File.Delete archivePath

    use archive = ZipFile.Open(archivePath, ZipArchiveMode.Create)
    archive.CreateEntryFromFile(outDir @@ exeName, exeName) |> ignore
    Trace.tracefn "Wrote %s" archivePath

Target.create "PackMcpSelfContained" (fun _ ->
    let version = localProjectVersion mcpProj

    for rid in selfContainedRids do
        publishSelfContained rid
        archiveSelfContained version rid)

// A third distribution channel again, this time via MCPB (github.com/modelcontextprotocol/
// mcpb, "MCP Bundles" - formerly DXT) - Claude for macOS/Windows's own one-click local
// server install mechanism, no command line at all needed by the end user. Verified this is
// real and not just the registry's own `mcpb` package-type keyword: fetched the actual spec
// (MANIFEST.md) and its official CLI (`npm install -g @anthropic-ai/mcpb`), hand-built a
// manifest.json + server/<exe> bundle for win-x64, ran it through the real `mcpb validate`/
// `pack`/`unpack`/`info` commands, then actually executed the unpacked binary (both the CLI
// `convert` path and the real stdio MCP server startup) to confirm the round trip produces
// something that actually runs, not just something the packer accepts.
//
// The `command`/`entry_point` path deliberately omits the .exe extension even for Windows
// RIDs - MANIFEST.md's own words: "For binaries, apps will automatically append .exe on
// Windows." The staged file itself still needs the real extension (mcpExeName), only the
// manifest's own reference to it doesn't.
let private mcpbCommandPath = "server/Kookerella.FsOpenXmlDsl.Mcp"

let private mcpbPlatform (rid: string) =
    if rid.StartsWith "win" then "win32"
    elif rid.StartsWith "osx" then "darwin"
    else "linux"

/// Reads the one field manifest.json needs that already has a canonical source elsewhere -
/// server.json's own "description" - rather than hand-typing a third copy of this string
/// that could drift from the other two the same way CLAUDE.md's "Keep these in sync" table
/// already warns about for the other copies.
let private mcpServerDescription () =
    use doc = JsonDocument.Parse(System.IO.File.ReadAllText mcpServerJson)
    doc.RootElement.GetProperty("description").GetString()

/// Built via JsonObject/JsonSerializer rather than a hand-formatted string template, so a
/// description (or any future field) containing a `"` or `\` can't corrupt the JSON - a real
/// risk a plain sprintf template would carry silently until the exact day some field's
/// content happened to need escaping.
let private mcpbManifestJson (version: string) (rid: string) : string =
    let author = JsonObject()
    author["name"] <- JsonValue.Create "Kookerella"

    let repository = JsonObject()
    repository["type"] <- JsonValue.Create "git"
    repository["url"] <- JsonValue.Create "https://github.com/Kookerella-Ltd/Kookerella.FsOpenXmlDsl"

    let mcpConfig = JsonObject()
    mcpConfig["command"] <- JsonValue.Create mcpbCommandPath
    mcpConfig["args"] <- JsonArray()
    mcpConfig["env"] <- JsonObject()

    let server = JsonObject()
    server["type"] <- JsonValue.Create "binary"
    server["entry_point"] <- JsonValue.Create mcpbCommandPath
    server["mcp_config"] <- mcpConfig

    let compatibility = JsonObject()
    let platforms = JsonArray()
    platforms.Add(JsonValue.Create(mcpbPlatform rid))
    compatibility["platforms"] <- platforms

    let manifest = JsonObject()
    manifest["manifest_version"] <- JsonValue.Create "0.3"
    manifest["name"] <- JsonValue.Create "fsopenxmldsl-mcp"
    manifest["display_name"] <- JsonValue.Create "Excel MCP Server (FsOpenXmlDsl)"
    manifest["version"] <- JsonValue.Create version
    manifest["description"] <- JsonValue.Create(mcpServerDescription ())
    manifest["author"] <- author
    manifest["repository"] <- repository
    manifest["homepage"] <-
        JsonValue.Create "https://github.com/Kookerella-Ltd/Kookerella.FsOpenXmlDsl/tree/master/src/Kookerella.FsOpenXmlDsl.Mcp"
    manifest["license"] <- JsonValue.Create "MIT"
    manifest["server"] <- server
    manifest["compatibility"] <- compatibility

    manifest.ToJsonString(JsonSerializerOptions(WriteIndented = true))

/// npm-installed CLI shims are `.cmd` files on Windows - confirmed empirically that
/// `Process.Start("mcpb", ...)` with `UseShellExecute = false` fails outright there ("the
/// system cannot find the file specified"), since it doesn't try PATHEXT resolution the way
/// a real shell does. Routing through `cmd.exe /c` fixes it there; plain invocation
/// elsewhere, where npm installs a real executable/shebang script instead.
let private runProcess (fileName: string) (args: string list) (workingDir: string) =
    let psi =
        if System.OperatingSystem.IsWindows() then
            let p = System.Diagnostics.ProcessStartInfo("cmd.exe")
            p.ArgumentList.Add("/c")
            p.ArgumentList.Add(fileName)
            for a in args do
                p.ArgumentList.Add(a)

            p
        else
            let p = System.Diagnostics.ProcessStartInfo(fileName)

            for a in args do
                p.ArgumentList.Add(a)

            p

    psi.WorkingDirectory <- workingDir
    psi.RedirectStandardOutput <- true
    psi.RedirectStandardError <- true
    psi.UseShellExecute <- false

    use proc =
        try
            System.Diagnostics.Process.Start(psi)
        with _ ->
            failwithf
                "Couldn't start '%s' - is it installed and on PATH? (For mcpb: npm install -g @anthropic-ai/mcpb)"
                fileName

    let stdout = proc.StandardOutput.ReadToEnd()
    let stderr = proc.StandardError.ReadToEnd()
    proc.WaitForExit()

    if stdout <> "" then
        Trace.log stdout

    if stderr <> "" then
        Trace.log stderr

    if proc.ExitCode <> 0 then
        failwithf "%s %s failed (exit %d) in %s - see output above." fileName (String.concat " " args) proc.ExitCode workingDir

/// Stages manifest.json + server/<exe> in a throwaway directory, then shells out to the
/// real `mcpb pack` CLI rather than hand-rolling the zip - not for the executable-bit
/// reason that seemed plausible at first (checked and ruled out, see archiveSelfContained's
/// own comment), but because MCPB is an actively evolving external spec (already at 0.3,
/// with a 0.4 revision documented) - shelling out to the authoritative tool means this
/// target stays correct as that spec changes, instead of this script silently drifting from
/// it the way a hand-tracked reimplementation would.
let private packMcpb (version: string) (rid: string) =
    let stagingDir = selfContainedDir @@ "mcpb-staging" @@ rid
    let serverDir = stagingDir @@ "server"
    System.IO.Directory.CreateDirectory(serverDir) |> ignore

    let exeName = mcpExeName rid
    System.IO.File.Copy(selfContainedDir @@ rid @@ exeName, serverDir @@ exeName, true)
    System.IO.File.WriteAllText(stagingDir @@ "manifest.json", mcpbManifestJson version rid)

    let outputPath = selfContainedDir @@ sprintf "fsopenxmldsl-mcp-%s-%s.mcpb" version rid

    if System.IO.File.Exists outputPath then
        System.IO.File.Delete outputPath

    runProcess "mcpb" [ "pack"; stagingDir; outputPath ] "."
    Trace.tracefn "Wrote %s" outputPath

Target.create "PackMcpMcpb" (fun _ ->
    let version = localProjectVersion mcpProj

    for rid in selfContainedRids do
        packMcpb version rid)

// Needs PackMcpSelfContained's own output (the per-RID single-file executables) to already
// exist - staging just copies from there rather than re-publishing.
"PackMcpSelfContained" ==> "PackMcpMcpb" |> ignore

"Clean" ==> "Restore" ==> "Build" ==> "TestFast" ==> "TestSlow" |> ignore

// Every Pack target depends on the full test gate, not just Build - so `fake build -t
// PushMcp` on its own still refuses to run against a failing suite, the same as the
// combined PublishAll target does.
"TestSlow" ==> "PackCore" ==> "PushCore" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackWrapper" ==> "PushWrapper" ==> "PublishAll" |> ignore
"TestSlow" ==> "PackMcp" ==> "PushMcp" ==> "PublishAll" |> ignore

// Deliberately NOT chained into PublishAll - uploading these as GitHub Release assets is
// its own separate, explicitly-triggered step (same reasoning as the MCP Registry sync
// needing a human login: this one needs a human decision about where these get hosted and
// when a new platform build is worth cutting, not something that should happen on every
// NuGet release automatically).
"TestSlow" ==> "PackMcpSelfContained" |> ignore

Target.runOrDefaultWithArguments "Build"
