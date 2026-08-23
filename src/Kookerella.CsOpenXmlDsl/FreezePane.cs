namespace Kookerella.CsOpenXmlDsl;

/// <summary>Number of leading rows/columns frozen in place, e.g. <c>new FreezePane(1, 0)</c>
/// freezes a header row - mirrors the F# core's <c>FreezePane</c>.</summary>
public sealed record FreezePane(int Rows, int Columns);
