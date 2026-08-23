# Kookerella.FsOpenXmlDsl.Mcp

<!-- mcp-name: io.github.MarkNicholls/fsopenxmldsl-mcp -->

An [MCP](https://modelcontextprotocol.io) (Model Context Protocol) server that exposes
`Kookerella.FsOpenXmlDsl`'s Excel read/write/code-generation capabilities as tools any
MCP-compatible AI agent can call directly - build a workbook, read one back, or regenerate
its F# source - without writing any F# itself.

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

## Scope

This is a deliberately narrow first pass over the library, not the whole thing: plain cell
values and formulas only, addressed by a row/column grid per sheet. Cell styling, tables,
charts, images, pivot tables, sparklines, VBA, conditional formatting, and everything else
`Kookerella.FsOpenXmlDsl` supports are not exposed as tools here — an agent that needs those
should reference the library directly (or use `generate_fsharp_script` on a file that
already has them, to see the F# for it). See the main project's
[MAPPING.md](https://github.com/Kookerella-Ltd/Kookerella.FsOpenXmlDsl/blob/master/MAPPING.md)
for the full picture of what the underlying library does and doesn't model.

`create_workbook`'s formula cells never carry a cached value (there's no formula
evaluation engine anywhere in this stack - see the main project's README for why that
matters for anything that isn't opened in real Excel first).
