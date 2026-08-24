namespace Kookerella.CsOpenXmlDsl;

/// <summary>A data validation rule applied over a range (a sheet can have several). Mirrors
/// the F# core's <c>DataValidationEntry</c>. Immutable - <see cref="WithAlert"/> returns a
/// new instance.</summary>
public sealed record DataValidationEntry(CellPosition TopLeft, CellPosition BottomRight, ValidationKind Kind)
{
    public ValidationAlert Alert { get; init; } = ValidationAlert.Default;

    public static DataValidationEntry Of(string topLeftA1, string bottomRightA1, ValidationKind kind) =>
        new(CellPosition.FromA1(topLeftA1), CellPosition.FromA1(bottomRightA1), kind);

    public DataValidationEntry WithAlert(ValidationAlert alert) => this with { Alert = alert };
}
