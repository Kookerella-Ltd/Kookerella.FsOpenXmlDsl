# Kookerella.FsOpenXmlDsl.Mcp

<!-- mcp-name: io.github.MarkNicholls/fsopenxmldsl-mcp -->

An [MCP](https://modelcontextprotocol.io) (Model Context Protocol) server that exposes
`Kookerella.FsOpenXmlDsl`'s Excel read/write/code-generation capabilities as tools any
MCP-compatible AI agent can call directly - build a workbook, read one back, or regenerate
its F#, C#, XML, or JSON representation - without writing any code itself.

**Most Excel libraries only go one direction**: build a workbook from scratch, or mutate an
existing one, through an imperative object model. This one also goes the other way - read
any existing `.xlsx`/`.xlsm` and hand back idiomatic, runnable F# or C# source (or plain
XML or JSON, each against a real schema) that rebuilds an equivalent file, and build a new
one from that XML/JSON directly. A decompiler for spreadsheets, not just a writer. That's
available three ways from this one binary: as MCP tools (`generate_fsharp_script`/
`generate_csharp_script`/`generate_xml`/`create_workbook_from_xml`/`generate_json`/
`create_workbook_from_json`) for an AI agent, as a plain CLI (`fsopenxmldsl-mcp
convert`/`build`) for anyone who isn't going through an MCP client, and as direct library
calls (`Workbook.generateScript`/`CsCodeGen.Generate`/`Xml.ofWorkbook`/`Xml.toWorkbook`/
`Json.ofWorkbook`/`Json.toWorkbook`) for either to call themselves.

The XML/JSON directions each have two concrete uses beyond code generation: **build an
`.xlsx` from XML/JSON a transform engine already produces** (an XSLT pipeline, or any
templating/generation script, can target Excel with no code at all), and **convert an
existing `.xlsx` to XML/JSON for version control** - `.xlsx` is a binary ZIP, so `git diff`
on one is useless, but `generate_xml`/`generate_json`'s output is deterministically ordered
(sorted by cell position, or by name for defined names) regardless of the source workbook's
own list order, so a real content change produces a small, isolated diff rather than a
spurious one from reshuffled rows.

This runs **locally**, as a subprocess your MCP client launches over stdio - there's no
hosted service, no network address, and no account to sign up for. It's distributed as a
[.NET tool](https://learn.microsoft.com/dotnet/core/tools/global-tools) for exactly that
reason: an MCP client just needs a command it can run, the same way it'd run any other CLI.

## Install

```bash
dotnet tool install -g Kookerella.FsOpenXmlDsl.Mcp
```

This installs the `fsopenxmldsl-mcp` command onto your PATH.

## Configure your MCP client

Point your client at the installed command. For example, in a client that reads a JSON
config with a `mcpServers` map:

```json
{
  "mcpServers": {
    "fsopenxmldsl": {
      "command": "fsopenxmldsl-mcp"
    }
  }
}
```

## Command-line usage

The same binary also works as a plain CLI, for converting a file without an MCP client at
all - `fsopenxmldsl-mcp` with no arguments starts the MCP server (as above); with a `convert`
or `build` first argument it runs once and exits:

```bash
fsopenxmldsl-mcp convert report.xlsx --lang csharp
```

Prints the equivalent C# source to stdout. Options:

- `--lang`/`-l` (required) — `fsharp`, `csharp`, `xml`, or `json`.
- `-o`/`--output <file>` — write the result to a file instead of stdout.
- `--rebuild-as <name.xlsx>` — `fsharp`/`csharp` only, the filename the *generated script
  itself* saves its rebuilt workbook to when run (default `output.xlsx`). Ignored for
  `--lang xml`/`json`, which have no script to embed a save path into.

```bash
fsopenxmldsl-mcp convert report.xlsx --lang fsharp -o report.fsx --rebuild-as rebuilt.xlsx
dotnet fsi report.fsx
```

`build` is the inverse of `convert --lang xml`/`convert --lang json` - it takes XML matching
`Xml.xsd` or JSON matching `Json.schema.json` and produces an `.xlsx` directly, for a caller
(e.g. an XSLT pipeline, or a plain JSON-emitting script) that already produces data that way
and wants to reach Excel without writing any code. Which format `build` reads is inferred
from the input file's own extension (`.xml` or `.json`):

```bash
fsopenxmldsl-mcp convert report.xlsx --lang xml -o report.xml   # .xlsx -> XML
fsopenxmldsl-mcp build report.xml rebuilt.xlsx                  # XML -> .xlsx

fsopenxmldsl-mcp convert report.xlsx --lang json -o report.json # .xlsx -> JSON
fsopenxmldsl-mcp build report.json rebuilt.xlsx                 # JSON -> .xlsx
```

## Tools

- **`create_workbook(path, sheets)`** — creates a new `.xlsx` from a simple grid of
  sheets/rows/cells. Each cell is given as plain text, the same way you'd type it into
  Excel: a leading `=` makes it a formula (e.g. `"=SUM(A1:A2)"`), `"true"`/`"false"` makes
  a boolean, a bare number is numeric, anything else is text.
- **`read_workbook(path)`** — reads an existing `.xlsx`/`.xlsm` and returns its sheets and
  cells as JSON, using the same text convention as `create_workbook`'s input (so a value
  round-trips through both tools unchanged).
- **`generate_fsharp_script(path, outputFileName)`** — reads an existing workbook and
  returns a self-contained F# script that rebuilds an equivalent file when run via
  `dotnet fsi`, using the full `Kookerella.FsOpenXmlDsl` API.
- **`generate_csharp_script(path, outputFileName)`** — the C# equivalent, for a caller who
  wants pasteable/runnable C# rather than F#. Reads an existing workbook through
  `Kookerella.CsOpenXmlDsl` (the idiomatic C# wrapper - now full feature parity with the F#
  core, see its own scope below) and returns a self-contained `.cs` file targeting .NET 10's
  file-based apps feature: `dotnet run <file>.cs`, no `.csproj` needed. References the
  published `Kookerella.CsOpenXmlDsl` NuGet package (via a `#:package` directive pinned to
  the version this server was built against), so the result runs on any machine with the
  .NET 10 SDK, not just this one.
- **`generate_xml(path)`** — a plain-data alternative to the two `generate_*_script` tools:
  reads an existing workbook and returns it as XML, validated against
  `Kookerella.FsOpenXmlDsl`'s own embedded schema (`Xml.xsd`). No `outputFileName` parameter,
  since the result is data, not a runnable script with a save path to embed.
- **`create_workbook_from_xml(xml, path)`** — the inverse of `generate_xml`: builds a new
  workbook from XML matching `Xml.xsd` and saves it to `path`. Unlike `create_workbook`,
  this isn't limited to plain cell values - styling, tables, charts, and every other
  modeled feature can be expressed in the XML, since it goes through the same schema
  `generate_xml` produces.
- **`generate_json(path)`** — the JSON equivalent of `generate_xml`: reads an existing
  workbook and returns it as JSON. Unlike `generate_xml`, there's no embedded runtime
  schema for this direction - `Json.schema.json` (in the main project's repo) is test-suite
  only, not a public API.
- **`create_workbook_from_json(json, path)`** — the inverse of `generate_json`: builds a new
  workbook from JSON matching the shape `generate_json` produces and saves it to `path`.
  Same relationship to `create_workbook` as `create_workbook_from_xml` has - not limited to
  plain cell values.

## Scope

`create_workbook`/`read_workbook` are a deliberately narrow first pass over the library, not
the whole thing: plain cell values and formulas only, addressed by a row/column grid per
sheet. Cell styling, tables, charts, images, pivot tables, sparklines, conditional
formatting, and everything else `Kookerella.FsOpenXmlDsl` supports aren't exposed through
those two tools — an agent that needs those should reference the library directly, or use
one of the other six tools on a file that already has them to see it represented as
source or data. All six cover the full worksheet/workbook-level feature set:
`generate_fsharp_script` the full F# core, `generate_csharp_script`/`generate_xml`/
`create_workbook_from_xml`/`generate_json`/`create_workbook_from_json` everything
`Kookerella.CsOpenXmlDsl`/`Xml.fs`/`Json.fs` model, which are now the same feature set as
the F# core. See the main project's
[MAPPING.md](https://github.com/Kookerella-Ltd/Kookerella.FsOpenXmlDsl/blob/master/MAPPING.md)
for the full picture of what the underlying library does and doesn't model.

`create_workbook`'s formula cells never carry a cached value (there's no formula
evaluation engine anywhere in this stack - see the main project's README for why that
matters for anything that isn't opened in real Excel first).
