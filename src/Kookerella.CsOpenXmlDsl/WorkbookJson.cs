using System.Text.Json.Nodes;
using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Converts a <see cref="Workbook"/> to/from JSON matching the shape documented by
/// <c>Json.schema.json</c> - a thin wrapper over <c>Kookerella.FsOpenXmlDsl.Json.ofWorkbook</c>/
/// <c>toWorkbook</c>, the JSON-side equivalent of <see cref="WorkbookXml"/>. Unlike the XML
/// side, nothing here validates against the schema at runtime - see the F# core's own
/// <c>Json.fs</c> doc comment for why (no .NET-built-in JSON Schema validator the way
/// <c>System.Xml.Schema</c> exists for XML, so wiring that up here would mean adding a
/// runtime dependency just for this).
/// </summary>
public static class WorkbookJson
{
    public static JsonObject ToJson(Workbook workbook) =>
        Fs.Json.ofWorkbook(WorkbookConverter.ToFSharp(workbook));

    public static Workbook FromJson(JsonObject json) =>
        WorkbookConverter.FromFSharp(Fs.Json.toWorkbook(json));
}
