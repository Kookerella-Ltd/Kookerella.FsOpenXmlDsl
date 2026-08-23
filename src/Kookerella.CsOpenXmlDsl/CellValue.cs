namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// A cell's content - a closed set of immutable cases, mirroring the F# core's own
/// <c>CellValue</c> discriminated union. C# has no true union type, so this uses the
/// standard "sealed hierarchy with a private base constructor" pattern: only the nested
/// records below can ever derive from it, so a <c>switch</c> over these five is exhaustive
/// in practice even though the compiler can't prove it. There is no <c>Empty</c> case here
/// (unlike the F# side) - simply don't add a cell with nothing in it.
/// </summary>
public abstract record CellValue
{
    private CellValue() { }

    /// <summary>Plain text. Always one uniformly-styled run - see the F# core's own
    /// <c>CellValue.Text</c> doc comment on why rich text runs aren't modeled.</summary>
    public sealed record Text(string Value) : CellValue;

    public sealed record Number(double Value) : CellValue;

    public sealed record Boolean(bool Value) : CellValue;

    public sealed record Date(DateTime Value) : CellValue;

    /// <summary><paramref name="Expression"/> excludes the leading "=". <paramref
    /// name="CachedValue"/> is what Excel shows before it recalculates on open; leave it
    /// <see langword="null"/> to force a recalculation on first open - see the main
    /// project's README for why that's only safe when a human opens the result in Excel
    /// first, since nothing in this stack evaluates formulas itself.</summary>
    public sealed record Formula(string Expression, double? CachedValue = null) : CellValue;
}
