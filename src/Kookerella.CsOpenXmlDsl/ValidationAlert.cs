namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// The non-essential parts of a data validation: whether blanks are allowed, and the
/// optional input prompt / error alert shown to the user. Kept separate from <see
/// cref="ValidationKind"/> so the common case (just a rule, no custom messages) doesn't need
/// to mention any of this. Mirrors the F# core's <c>ValidationAlert</c>. Immutable - every
/// <c>With*</c> method returns a new instance.
/// </summary>
public sealed record ValidationAlert
{
    public bool AllowBlank { get; init; } = true;
    public ErrorAlertStyle ErrorStyle { get; init; } = ErrorAlertStyle.Stop;
    public string? ErrorTitle { get; init; }
    public string? ErrorMessage { get; init; }
    public string? InputTitle { get; init; }
    public string? InputMessage { get; init; }

    public static readonly ValidationAlert Default = new();

    public ValidationAlert WithAllowBlank(bool allowBlank) => this with { AllowBlank = allowBlank };
    public ValidationAlert WithErrorStyle(ErrorAlertStyle style) => this with { ErrorStyle = style };
    public ValidationAlert WithErrorTitle(string title) => this with { ErrorTitle = title };
    public ValidationAlert WithErrorMessage(string message) => this with { ErrorMessage = message };
    public ValidationAlert WithInputTitle(string title) => this with { InputTitle = title };
    public ValidationAlert WithInputMessage(string message) => this with { InputMessage = message };
}
