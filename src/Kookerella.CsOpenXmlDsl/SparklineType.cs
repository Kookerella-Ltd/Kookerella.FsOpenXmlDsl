namespace Kookerella.CsOpenXmlDsl;

/// <summary>The visual shape of a sparkline group - mirrors the F# core's
/// <c>SparklineType</c>, Excel's own three kinds. <see cref="WinLoss"/> is what Excel's UI
/// calls "Win/Loss" (OOXML names the same enum value <c>stacked</c>).</summary>
public enum SparklineType
{
    Line,
    Column,
    WinLoss
}
