namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// One column of an Excel Table. <see cref="Name"/> must be unique within the table and,
/// per Excel's own requirement, should match the text of the header cell at that column's
/// position in the worksheet - this wrapper doesn't synthesize that header cell for you,
/// it only describes metadata layered on top of cells you've already placed. Mirrors the
/// F# core's <c>TableColumn</c>.
/// </summary>
public sealed record TableColumn(string Name, string? CalculatedFormula = null);
