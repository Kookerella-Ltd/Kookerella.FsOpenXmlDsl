namespace Kookerella.CsOpenXmlDsl;

/// <summary>Shared by conditional formatting's <see cref="ConditionalFormatRule.CellValueRule"/>
/// and (once exposed) data validation's numeric/text-length rules - OOXML uses the same
/// comparison vocabulary for both. Mirrors the F# core's <c>ComparisonOperator</c>.</summary>
public enum ComparisonOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,

    /// <summary>Needs both <see cref="ConditionalFormatRule.CellValueRule.Formula1"/> and
    /// <see cref="ConditionalFormatRule.CellValueRule.Formula2"/>.</summary>
    Between,

    /// <summary>Needs both <see cref="ConditionalFormatRule.CellValueRule.Formula1"/> and
    /// <see cref="ConditionalFormatRule.CellValueRule.Formula2"/>.</summary>
    NotBetween
}
