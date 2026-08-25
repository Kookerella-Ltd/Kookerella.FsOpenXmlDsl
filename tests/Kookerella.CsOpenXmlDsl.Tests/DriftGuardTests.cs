using Microsoft.FSharp.Reflection;
using Xunit;
using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl.Tests;

/// <summary>
/// A structural tripwire against the specific failure mode this project has already hit
/// more than once: a case gets added to an F# core discriminated union, and the C# wrapper
/// silently doesn't grow a matching case, because nothing forces anyone to notice. This
/// doesn't verify the C# side is *correct* - only that its case count hasn't fallen behind
/// the F# core's. See CLAUDE.md's "Adding a feature" checklist for what to actually do once
/// this test points at a gap.
///
/// Deliberately count-based rather than name-based: several mirrors rename cases on
/// purpose (F#'s <c>AlignLeft</c>/<c>GeneralAlign</c> etc. become plain <c>Left</c>/
/// <c>General</c> in <see cref="HorizontalCellAlignment"/>), so comparing exact case names
/// would false-positive on those. A count mismatch is a strictly weaker but far more
/// robust signal - it still catches an added-and-forgotten case, it just doesn't name which
/// one, which is an acceptable trade for not needing a name-translation table per type.
/// </summary>
public class DriftGuardTests
{
    /// <summary>F# cases with no C# counterpart on purpose, each backed by a real reason -
    /// not an oversight this test should ever "fix" by someone just adding to this list.
    /// Adding an entry here without a corresponding doc trail (a MAPPING.md gap, or a
    /// wrapper doc comment explaining the omission) defeats the point of this test.</summary>
    private static readonly Dictionary<Type, string[]> KnownGaps = new()
    {
        // BorderLineStyle.Other is an escape hatch for a raw OOXML style name Core doesn't
        // have a named case for - see BorderSide's own doc comment in WorkbookConverter.cs.
        [typeof(Fs.Styles.BorderLineStyle)] = ["Other"],
        // NumberFormat.Custom carries a raw format-code string, exposed separately via
        // CellStyle.CustomNumberFormat rather than as a NumberFormatKind member.
        [typeof(Fs.Styles.NumberFormat)] = ["Custom"],
        // CellValue.Empty has no CellValue.cs case at all - an empty cell is a null
        // CellValue/CellStyle reference in the C# model instead of an explicit case.
        [typeof(Fs.Model.CellValue)] = ["Empty"],
    };

    private static int FsCaseCount(Type fsUnionType)
    {
        var gapCount = KnownGaps.TryGetValue(fsUnionType, out var gaps) ? gaps.Length : 0;
        return FSharpType.GetUnionCases(fsUnionType, null).Length - gapCount;
    }

    private static int CsEnumCaseCount(Type csEnumType) => Enum.GetNames(csEnumType).Length;

    /// <summary>Counts a closed hierarchy's case types the same way the codebase builds
    /// them (nested sealed records directly under an abstract base) - see
    /// <c>ConditionalFormatRule.cs</c>/<c>CellValue.cs</c>/etc. for the pattern.</summary>
    private static int CsClosedHierarchyCaseCount(Type csAbstractBaseType) =>
        csAbstractBaseType.Assembly.GetTypes().Count(t => t.BaseType == csAbstractBaseType);

    public static IEnumerable<object[]> EnumMirrors =>
        new List<(Type Fs, Type Cs)>
        {
            (typeof(Fs.Styles.BorderLineStyle), typeof(BorderLineStyle)),
            (typeof(Fs.Styles.HorizontalAlignment), typeof(HorizontalCellAlignment)),
            (typeof(Fs.Styles.VerticalAlignment), typeof(VerticalCellAlignment)),
            (typeof(Fs.Styles.NumberFormat), typeof(NumberFormatKind)),
            (typeof(Fs.ChartType), typeof(ChartType)),
            (typeof(Fs.PivotAggregation), typeof(PivotAggregation)),
            (typeof(Fs.SparklineType), typeof(SparklineType)),
            (typeof(Fs.ComparisonOperator), typeof(ComparisonOperator)),
            (typeof(Fs.ErrorAlertStyle), typeof(ErrorAlertStyle)),
            (typeof(Fs.PageOrientation), typeof(PageOrientation)),
            (typeof(Fs.ImageFormat), typeof(ImageFormat))
        }.Select(pair => new object[] { pair.Fs, pair.Cs });

    [Theory]
    [MemberData(nameof(EnumMirrors))]
    public void Fs_union_case_count_matches_Cs_enum(Type fsType, Type csType)
    {
        Assert.True(
            FsCaseCount(fsType) == CsEnumCaseCount(csType),
            $"{fsType.FullName} has {FsCaseCount(fsType)} case(s) after known gaps, but " +
            $"{csType.FullName} has {CsEnumCaseCount(csType)} enum member(s). A case was " +
            "likely added to one side without the other - see CLAUDE.md's 'Adding a " +
            "feature' checklist, or add a documented entry to this test's KnownGaps if the " +
            "omission is deliberate.");
    }

    public static IEnumerable<object[]> ClosedHierarchyMirrors =>
        new List<(Type Fs, Type Cs)>
        {
            (typeof(Fs.Model.CellValue), typeof(CellValue)),
            (typeof(Fs.ConditionalFormatRule), typeof(ConditionalFormatRule)),
            (typeof(Fs.ValidationKind), typeof(ValidationKind)),
            (typeof(Fs.HyperlinkTarget), typeof(HyperlinkTarget)),
            (typeof(Fs.DefinedNameScope), typeof(DefinedNameScope)),
            (typeof(Fs.PaperSize), typeof(PaperSize)),
            (typeof(Fs.PrintScaling), typeof(PrintScaling))
        }.Select(pair => new object[] { pair.Fs, pair.Cs });

    [Theory]
    [MemberData(nameof(ClosedHierarchyMirrors))]
    public void Fs_union_case_count_matches_Cs_closed_hierarchy(Type fsType, Type csType)
    {
        Assert.True(
            FsCaseCount(fsType) == CsClosedHierarchyCaseCount(csType),
            $"{fsType.FullName} has {FsCaseCount(fsType)} case(s) after known gaps, but " +
            $"{csType.FullName} has {CsClosedHierarchyCaseCount(csType)} nested case " +
            "type(s). A case was likely added to one side without the other - see " +
            "CLAUDE.md's 'Adding a feature' checklist, or add a documented entry to this " +
            "test's KnownGaps if the omission is deliberate.");
    }
}
