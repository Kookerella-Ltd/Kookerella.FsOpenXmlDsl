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
   `Interpreter/CodeGen.fs` (F# source-gen) + `Xml.fs`/`Xml.xsd` (the XML surface - a new
   field/case needs a schema change too, not just a code change; the "every generated
   `workbook.xml` validates against `Xml.xsd`" check inside `verifyScenarioNamed` is what
   catches the two drifting apart) + `Json.fs`/`Json.schema.json` (the JSON surface - same
   feature set as `Xml.fs`, same drift risk; `verifyScenarioNamed` also writes and validates
   an `Examples/<name>/workbook.json` for every scenario, the same way it does
   `workbook.xml` - `Json.schema.json` is test-suite only, so this check runs from
   `JsonTests.assertJsonSchemaValid`, not a public `Json.schemaSet()`-equivalent API - see
   `Json.fs`'s own doc comment) + a test in `tests/Kookerella.FsOpenXmlDsl.Tests` for both
   `XmlTests.fs` and `JsonTests.fs`, and the F# core's own `<Description>` in
   `Kookerella.FsOpenXmlDsl.fsproj` (this package went
   stale once already - `Xml.fs` shipped in source for several commits before anyone
   noticed the *published* package was still an old version without it, because the Mcp
   server builds against local source via `ProjectReference` and never surfaced the gap).
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
   `dotnet fake run build.fsx -t PublishAll` (see `build.fsx` — this runs the full test gate
   first). **Use `PublishAll`, never `Push<Core|Wrapper|Mcp>` individually, and never a
   manual `dotnet pack`/`dotnet nuget push`** - two real incidents behind this, not caution
   for its own sake:
   - `build.fsx` exists specifically so a release can't skip the test gate, but that only
     holds if it's actually invoked; one release went out via plain `dotnet pack` anyway,
     purely out of habit, defeating the point.
   - The C# wrapper's `ProjectReference` to the F# core gets converted into a NuGet
     dependency *floor* (`>= x.y.z`) at whatever version the core happened to be at pack
     time - `dotnet pack`/NuGet do this correctly and automatically, but only the wrapper's
     *own* next pack captures a newer core version. Calling `PushCore` alone (three times in
     a row, across the Xml.fs/Json.fs/font-ordering-fix releases) left the *published*
     wrapper silently declaring a dependency floor two minor versions behind the core for
     weeks, invisible locally because a local checkout is always self-consistent by
     construction - it only showed up when a from-scratch NuGet restore (a demo project,
     the decompiled `.g.cs`'s own `#:package` restore) pulled the ancient floor version.
     `push` (in `build.fsx`) is deliberately a no-op for whichever package(s) didn't change
     this release rather than erroring on "already published", specifically so `PublishAll`
     is always safe and always the right thing to run — every release, not just ones that
     "feel like" they touched more than one package. `PublishAll`'s own action (not a
     separate target merely chained after it, which would only run on `-t
     VerifyDependencyFreshness`, never on the `-t PublishAll` people actually type) then
     checks nuget.org's own live state and fails loudly if the wrapper's published floor
     still doesn't match the core's latest published version - a backstop for exactly the
     mistake above, in case `PublishAll` gets bypassed anyway. Also independently invocable
     any time via `dotnet fake run build.fsx -t VerifyDependencyFreshness --single-target`.
   - The Mcp tool has the same problem one layer deeper, discovered live while building an
     unrelated demo project: it's a self-contained `dotnet tool` (`PackAsTool`) that bundles
     its full dependency closure as plain files rather than declaring a NuGet floor, built
     via its own `ProjectReference`s the same way - so it goes stale the same way, for the
     same reason, and `dotnet tool update` can't fix an already-stale *published* package,
     only install whatever's currently there. `PublishAll`'s action also downloads the live
     Mcp `.nupkg` and reads its bundled `Kookerella.FsOpenXmlDsl.dll`/
     `Kookerella.CsOpenXmlDsl.dll` version metadata directly, comparing against the core's/
     wrapper's own latest published versions.
   NuGet indexing typically takes 5–20 minutes after a successful push before the new
   version resolves anywhere (search, flatcontainer index, `dotnet restore`) — don't assume
   a push failed just because it isn't visible yet; check nuget.org's own package page (it
   updates before the API indexes do) before retrying.
6. **MCP Registry sync**, only if the Mcp package's version changed: `mcp-publisher login
   github` (interactive GitHub device-flow — this needs the user, it can't be scripted or
   run non-interactively) immediately followed by `mcp-publisher publish` from
   `src/Kookerella.FsOpenXmlDsl.Mcp/.mcp/` — the registry JWT is short-lived, so log in
   again right before publishing rather than reusing an older session. The registry
   publish will itself reject the request with a clear error if the NuGet version it
   references isn't indexed yet — that's the signal to wait, not a real failure.

## Keep these in sync

The checklist above is organized by "when adding a feature, touch these in order." This is
the same information reorganized as a flat list of every file/field that describes a
capability rather than implementing one - useful as a quick scan for "did I miss
something" regardless of what kind of change is in flight, since a doc/metadata-only change
(no new DSL feature at all) can still make one of these stale on its own.

[`docs.deps.json`](docs.deps.json) is the same checklist again, a third way: a structured
`trigger -> targets` list rather than prose, meant to be grepped/read by an LLM session at
the start of a change rather than re-derived from memory each time. It doesn't replace this
table or the checklist above - it's a query-friendly index into them, and only as trustworthy
as it's kept current (add an edge when something drifts for real, the same rule everything
else here follows - don't add one speculatively).

| File | What must stay accurate | Checked by |
|---|---|---|
| `README.md` (root) | Top summary, per-feature sections, Layout list | Nothing automated - read it |
| `llms.txt` | Top summary, per-package scope sections | Nothing automated - read it |
| `src/Kookerella.CsOpenXmlDsl/README.md` | Feature list, `## Scope` | Nothing automated |
| `src/Kookerella.FsOpenXmlDsl.Mcp/README.md` | Tool list, `## Scope`, CLI usage | Nothing automated |
| `src/Kookerella.FsOpenXmlDsl/Kookerella.FsOpenXmlDsl.fsproj` `<Description>` | Matches the F# core's actual feature set | Nothing automated - it's NuGet metadata, not code |
| `src/Kookerella.CsOpenXmlDsl/Kookerella.CsOpenXmlDsl.csproj` `<Description>` | Matches the C# wrapper's actual feature set | Nothing automated |
| `src/Kookerella.FsOpenXmlDsl.Mcp/Kookerella.FsOpenXmlDsl.Mcp.fsproj` `<Description>` | Matches the Mcp server's actual tool/CLI surface | Nothing automated |
| `src/Kookerella.FsOpenXmlDsl.Mcp/.mcp/server.json` | `description` (≤100 chars, registry-enforced) **and** both `version` fields (top-level and `packages[].version`) match the `.fsproj`'s `<Version>` | The registry publish itself rejects a version it can't find on NuGet, and `VerifyDependencyFreshness`/`PublishAll`'s own action (`verifyServerJsonVersion`) checks both version fields against the `.fsproj` - but nothing checks `description` accuracy |
| `src/Kookerella.FsOpenXmlDsl.Mcp/WorkbookTools.fs` | Every tool's `[<Description>]` text, and the `WorkbookTools` type's own doc comment | Nothing automated - this is what an MCP client/agent actually reads, separate from any README |
| `src/Kookerella.FsOpenXmlDsl.Mcp/Dockerfile` | `COPY` list mirrors every `ProjectReference` in the `.fsproj` exactly | Nothing automated unless someone actually runs `docker build` |
| `src/Kookerella.FsOpenXmlDsl/Xml.xsd` | Matches what `Xml.fs`'s `ofWorkbook`/`toWorkbook` actually read and write | `assertXmlSchemaValid` inside `verifyScenarioNamed` - real, but only as strong as the scenarios that exist |
| Published `Kookerella.CsOpenXmlDsl`'s NuGet dependency floor on `Kookerella.FsOpenXmlDsl` | Must equal the F# core's latest *published* version, not just whatever's in the local `.fsproj` | `VerifyDependencyFreshness`/`PublishAll`'s own action in `build.fsx` (`verifyWrapperDependencyFloor`) - queries nuget.org's live state directly; only catches it if `PublishAll` (not a standalone `Push*`) is actually what ran |
| Published `Kookerella.FsOpenXmlDsl.Mcp`'s *bundled* `Kookerella.FsOpenXmlDsl.dll`/`Kookerella.CsOpenXmlDsl.dll` (it's a self-contained `dotnet tool`, no nuspec dependency floor to check instead) | Their assembly versions must equal the core's/wrapper's latest *published* versions | `VerifyDependencyFreshness`/`PublishAll`'s own action in `build.fsx` (`verifyMcpBundleFreshness`) - downloads the live `.nupkg` and reads the bundled DLLs' own version metadata; found live and stale (bundling a pre-font-fix core) the first time this was built, from a from-scratch `dotnet tool update` + demo project restore, not from anything local |
| `src/Kookerella.FsOpenXmlDsl/Json.schema.json` | Matches what `Json.fs`'s `ofWorkbook`/`toWorkbook` actually read and write | `assertJsonSchemaValid`, called from every `JsonTests.fs` round trip via its `roundTrip` helper, **and** from `verifyScenarioNamed` for every `Examples/` scenario (writes `workbook.json`, same as `workbook.xml`/`Xml.xsd`) - real, but only as strong as the scenarios that exist |
| `tests/Kookerella.CsOpenXmlDsl.Tests/ExampleTests.cs`'s `AssertWorkbooksMatch` | Checks every `Sheet`/`Workbook` field the slow round-trip theory is supposed to verify | Nothing - it silently stopped covering five-plus features in a row once already |
| `tests/Kookerella.CsOpenXmlDsl.Tests/DriftGuardTests.cs`'s `EnumMirrors`/`ClosedHierarchyMirrors` | Lists every F# DU mirrored as a C# enum/closed hierarchy | Only guards types already registered with it - a brand-new type is invisible until added |
| `MAPPING.md` | What the F# core maps 1:1 vs. approximates vs. doesn't model | Nothing automated - only touch this for a new OOXML-level capability, not a wrapper-level one |

Three packages, three `.fsproj`/`.csproj` `<Version>` fields, one `server.json` with two
more copies of one of those three - a version bump in one place and not the others is the
single most common way this list goes stale. When in doubt, grep the whole repo for the old
version string before considering a bump finished.

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
  `PushCore`/`PushWrapper`/`PushMcp`, `PublishAll`, `PackMcpSelfContained`).
- `PackMcpSelfContained` produces a self-contained, single-file build of the Mcp server per
  platform (win/linux/osx × x64/arm64) - a second distribution channel alongside the
  `dotnet tool` package, for a machine with no .NET runtime pre-installed at all. Native AOT
  (smaller, faster) was tried and rejected empirically, not just assumed impractical: it
  compiles and produces a native binary, but crashes at runtime the first time it hits any
  `sprintf`/`printf`/`failwithf` call - F#'s own formatting machinery is reflection-based
  (`MakeGenericMethod`) in a way Native AOT's trimmer can't resolve statically, and those
  three are used pervasively throughout ordinary F# code, not one isolated call site. Not
  wired into `PublishAll` - uploading these as release assets is a separate, manual step
  (same reasoning as the MCP Registry sync needing a human login).
- Fast tests only: `dotnet test --filter "Category!=Slow"`. The slow group actually
  executes each generated example script via `dotnet run`/`dotnet fsi` and diffs the
  result against the committed file — always run this before any release, not just fast
  tests.
