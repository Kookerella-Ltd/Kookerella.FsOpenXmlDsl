namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Visual style reference for a Table - a named built-in or custom table style plus the
/// banded-rows/-columns and first/last-column-emphasis toggles Excel's Table Design ribbon
/// exposes. <see cref="Name"/> is a raw style name (e.g. <c>"TableStyleMedium2"</c>, one of
/// Excel's several dozen built-ins) - style *definitions* (their colors/fonts) aren't
/// modeled, only the reference to one. Mirrors the F# core's <c>TableStyle</c>. Immutable -
/// every <c>With*</c>/<c>Without*</c> method returns a new instance.
/// </summary>
public sealed record TableStyle
{
    public string? Name { get; init; } = "TableStyleMedium2";
    public bool ShowFirstColumn { get; init; }
    public bool ShowLastColumn { get; init; }
    public bool ShowRowStripes { get; init; } = true;
    public bool ShowColumnStripes { get; init; }

    /// <summary>Matches what Excel's own "Insert &gt; Table" applies by default.</summary>
    public static readonly TableStyle Default = new();

    public TableStyle WithName(string? name) => this with { Name = name };
    public TableStyle WithFirstColumnEmphasis() => this with { ShowFirstColumn = true };
    public TableStyle WithLastColumnEmphasis() => this with { ShowLastColumn = true };
    public TableStyle WithoutRowStripes() => this with { ShowRowStripes = false };
    public TableStyle WithColumnStripes() => this with { ShowColumnStripes = true };
}
