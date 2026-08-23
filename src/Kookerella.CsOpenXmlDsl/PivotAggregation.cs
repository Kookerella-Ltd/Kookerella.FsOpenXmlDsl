namespace Kookerella.CsOpenXmlDsl;

/// <summary>How a pivot table's value field is aggregated across each group of matching
/// source rows. Mirrors the F# core's <c>PivotAggregation</c>: covers the six most commonly
/// used consolidation functions (Product/StdDev/StdDevP/Var/VarP aren't modeled in either
/// layer).</summary>
public enum PivotAggregation
{
    Sum,
    Count,
    CountNumbers,
    Average,
    Min,
    Max
}
