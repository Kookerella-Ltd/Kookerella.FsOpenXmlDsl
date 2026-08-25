# CLAUDE.md

Instructions for any Claude Code session working in this repo. This file exists because
every item on the checklist below has drifted out of sync at least once in this project's
history — each was found reactively (a stray user question, or noticing it while checking
something unrelated), never automatically. Treat this as the fix for that: a fact-of-life
checklist, not aspirational.

## Repo layout

Three packages from one repo:
- `src/Kookerella.FsOpenXmlDsl` — the F# core: a typesafe DSL over the OOXML SpreadsheetML
  schema, interpreted by `Interpreter/Writer.fs` and reversed by `Interpreter/Reader.fs`.
- `src/Kookerella.CsOpenXmlDsl` — an idiomatic, immutable, fluent C# wrapper over the F#
  core. As of this writing it has full feature parity with the F# core at the
  worksheet/workbook level.
- `src/Kookerella.FsOpenXmlDsl.Mcp` — a local MCP server exposing both as agent tools, and
  as a plain `fsopenxmldsl-mcp convert` CLI command.

## Adding a feature to the DSL — checklist

When a new type/DU/field is added to the F# core's model (`Worksheet`/`Workbook` or
anything they contain), work through **all** of these, in order. Skipping one is exactly
how past drift happened — each step below has a real, specific incident behind it.

1. **F# core** (`src/Kookerella.FsOpenXmlDsl`): model + `Writer.fs` + `Reader.fs` +
   `Interpreter/CodeGen.fs` (F# source-gen) + a test in
   `tests/Kookerella.FsOpenXmlDsl.Tests`.
2. **C# wrapper** (`src/Kookerella.CsOpenXmlDsl`), in this order:
   - Verify the F#/C# compiled interop shape via a `dotnet fsi` reflection scratch script
     **before** writing any conversion code — constructor parameter order, `NewCaseName`
     static factories on DU cases, exact property names. Never guess this.
   - New C# type file(s) mirroring the F# shape (parameterless-only DU → plain enum;
     data-carrying DU → closed hierarchy, mirroring `CellValue`'s pattern).
   - `Sheet.cs`/`Workbook.cs` — new property + fluent `With*`/`Add*` methods.
   - `WorkbookConverter.cs` — `ToFsX`/`FromFsX`, both directions.
   - `CsCodeGen.cs` — **before** writing tests, not after (a real gap shipped once because
     this was left until "later").
   - Tests: `Examples/` scenario(s) in `ExampleTests.cs` mirroring the F# test suite's
     *exact* reference data, unit tests in `WorkbookTests.cs`, **and** update
     `AssertWorkbooksMatch` in `ExampleTests.cs`. This one is easy to miss because the test
     suite still passes without it — it silently stopped checking round-trip fidelity for
     five-plus features in a row before anyone noticed.
   - If the new feature is a brand-new F# DU mirrored as a C# enum or closed hierarchy
     (not just a new case on an existing one), add it to `DriftGuardTests.cs`'s
     `EnumMirrors`/`ClosedHierarchyMirrors` list — that test only guards types already
     registered with it, so a new type is invisible to it until added.
   - The C# wrapper's own `README.md` (feature section + `## Scope`) and the `<Description>`
     in `Kookerella.CsOpenXmlDsl.csproj` (the NuGet "sales pitch" text — easy to forget
     since it's metadata, not code, and it directly shapes whether someone chooses this
     package at all).
3. **Root docs**:
   - Root `README.md` — it asserts the wrapper's scope in at least one place; don't assume
     old wording ("narrow first pass", feature lists) is still accurate.
   - `llms.txt` — the C# wrapper's "v1 scope" list and any "not exposed" list.
   - `MAPPING.md` only if this is a new OOXML-level capability in the F# core itself (not a
     C# wrapper concern).
4. **MCP server** (`src/Kookerella.FsOpenXmlDsl.Mcp`), if a `ProjectReference` or a
   referenced package's version changed:
   - `Dockerfile`'s `COPY` list must mirror every `ProjectReference` in the `.fsproj`
     exactly — it silently broke once when `Kookerella.CsOpenXmlDsl` was added as a
     reference and the Dockerfile wasn't touched; nothing catches this except actually
     running `docker build`.
   - Any MCP tool's `[<Description>]` text that asserts a feature scope (e.g.
     `generate_csharp_script`'s own doc string in `WorkbookTools.fs`) — this is what an MCP
     client/agent sees directly, separately from any README.
   - The Mcp project's own `README.md` and `.mcp/server.json`'s `description` field (note:
     the registry enforces a **100-character limit** on that field — it will reject a
     longer one at publish time, not at edit time).
5. **Version bump + release** — bump `<Version>` by hand in whichever project(s) changed
   (a semver judgment call, not automated), then run
   `dotnet fake run build.fsx -t Push<Core|Wrapper|Mcp>` (see `build.fsx` — this runs the
   full test gate first, even for a single Push target). NuGet indexing typically takes
   5–20 minutes after a successful push before the new version resolves anywhere (search,
   flatcontainer index, `dotnet restore`) — don't assume a push failed just because it
   isn't visible yet; check nuget.org's own package page (it updates before the API
   indexes do) before retrying.
6. **MCP Registry sync**, only if the Mcp package's version changed: `mcp-publisher login
   github` (interactive GitHub device-flow — this needs the user, it can't be scripted or
   run non-interactively) immediately followed by `mcp-publisher publish` from
   `src/Kookerella.FsOpenXmlDsl.Mcp/.mcp/` — the registry JWT is short-lived, so log in
   again right before publishing rather than reusing an older session. The registry
   publish will itself reject the request with a clear error if the NuGet version it
   references isn't indexed yet — that's the signal to wait, not a real failure.

## Process discipline

- Never commit or push without the user explicitly saying so in the current turn — a prior
  approval doesn't carry forward to later, unrelated changes.
- Never let a secret (API key, token) reach a command wrapper that logs its own full
  argument list — FAKE's `DotNet.exec` does this on every invocation, success or failure.
  Shell out via a raw process call with a redacted log line instead (see `build.fsx`'s
  `push` function for the pattern). Verify this by actually running the command and
  grepping captured output for the secret, not by reading the code and assuming it's safe.
- Verify, don't assume: run the actual build/test/tool before reporting something works.
  Several bugs on record here (an `#r` package-manager restriction, an FSharp.Core version
  conflict a build script hit, a Docker copy-list gap) were only found because something
  was actually executed, not because the code looked right.

## Build

- `dotnet tool restore` once (restores `fake-cli` from `.config/dotnet-tools.json`).
- `dotnet fake run build.fsx -t <Target>` — see `build.fsx` for the full target list
  (`Clean`, `Restore`, `Build`, `TestFast`, `TestSlow`, `PackCore`/`PackWrapper`/`PackMcp`,
  `PushCore`/`PushWrapper`/`PushMcp`, `PublishAll`).
- Fast tests only: `dotnet test --filter "Category!=Slow"`. The slow group actually
  executes each generated example script via `dotnet run`/`dotnet fsi` and diffs the
  result against the committed file — always run this before any release, not just fast
  tests.
