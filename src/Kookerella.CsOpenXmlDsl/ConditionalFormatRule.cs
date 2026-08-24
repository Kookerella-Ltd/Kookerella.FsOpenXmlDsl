namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A conditional formatting rule to apply over a range - a closed set of immutable cases,
/// mirroring the F# core's own <c>ConditionalFormatRule</c> discriminated union (same
/// "sealed hierarchy with a private base constructor" pattern <see cref="CellValue"/> uses).
/// <see cref="CellValueRule.Formula1"/>/<see cref="CellValueRule.Formula2"/> and <see
/// cref="FormulaRule.Formula"/> are raw formula text (same convention as <see
/// cref="CellValue.Formula"/>) - for <see cref="CellValueRule"/> these are literal values or
/// cell references compared against, not <c>=</c>-prefixed formulas.
/// <para>
/// Covers the common cases; icon sets, "top/bottom N", and the text/blank/error-contains
/// rule kinds aren't modeled - reference the F# core directly for those. Named <see
/// cref="ColorScale2"/>/<see cref="ColorScale3"/>/<see cref="DataBarRule"/> rather than bare
/// <c>ColorScale</c>/<c>DataBar</c> for the same reason the F# core does: avoiding a
/// collision with the OpenXml SDK's own identically-named types.
/// </para>
/// </summary>
public abstract record ConditionalFormatRule
{
    private ConditionalFormatRule() { }

    public sealed record CellValueRule(ComparisonOperator Operator, string Formula1, string? Formula2, CellStyle Style) : ConditionalFormatRule;

    public sealed record FormulaRule(string Formula, CellStyle Style) : ConditionalFormatRule;

    public sealed record ColorScale2(RgbColor MinColor, RgbColor MaxColor) : ConditionalFormatRule;

    public sealed record ColorScale3(RgbColor MinColor, RgbColor MidColor, RgbColor MaxColor) : ConditionalFormatRule;

    /// <summary>A single-color data bar with Excel's default automatic min/max thresholds.</summary>
    public sealed record DataBarRule(RgbColor Color) : ConditionalFormatRule;

    public sealed record DuplicateValuesRule(CellStyle Style) : ConditionalFormatRule;

    public sealed record UniqueValuesRule(CellStyle Style) : ConditionalFormatRule;
}
