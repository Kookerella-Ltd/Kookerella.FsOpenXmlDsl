namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// What kind of value a data validation rule accepts - a closed set of immutable cases,
/// mirroring the F# core's own <c>ValidationKind</c> discriminated union (same "sealed
/// hierarchy with a private base constructor" pattern <see cref="CellValue"/> uses). <see
/// cref="WholeNumberValidation.Formula1"/>/<c>Formula2</c> and friends are raw formula text,
/// the same convention as <see cref="ConditionalFormatRule"/>.
/// <para>
/// Covers the common cases; <c>Date</c>/<c>Time</c> validation and cross-sheet named-range
/// list sources aren't modeled - reference the F# core directly for those.
/// </para>
/// </summary>
public abstract record ValidationKind
{
    private ValidationKind() { }

    /// <summary>A fixed, inline dropdown list of choices.</summary>
    public sealed record ListValidation : ValidationKind
    {
        public IReadOnlyList<string> Items { get; }

        public ListValidation(params string[] items) => Items = items;
    }

    /// <summary>A dropdown list sourced from another range's values.</summary>
    public sealed record ListFromRangeValidation(CellPosition TopLeft, CellPosition BottomRight) : ValidationKind;

    public sealed record WholeNumberValidation(ComparisonOperator Operator, string Formula1, string? Formula2) : ValidationKind;

    public sealed record DecimalValidation(ComparisonOperator Operator, string Formula1, string? Formula2) : ValidationKind;

    public sealed record TextLengthValidation(ComparisonOperator Operator, string Formula1, string? Formula2) : ValidationKind;

    /// <summary>An arbitrary boolean formula.</summary>
    public sealed record CustomValidation(string Formula) : ValidationKind;
}
