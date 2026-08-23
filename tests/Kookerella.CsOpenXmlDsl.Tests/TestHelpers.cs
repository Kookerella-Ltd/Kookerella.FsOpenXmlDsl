using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Kookerella.CsOpenXmlDsl.Tests;

/// <summary>Shared fixtures/assertions used by both <c>WorkbookTests</c> (behavioral/edge-case
/// tests against throwaway temp files) and <c>ExampleTests</c> (feature-demonstration
/// scenarios saved into the checked-in <c>Examples/</c> folder) - pulled in via <c>using
/// static</c> in each so call sites read exactly as they did before this split.</summary>
internal static class TestHelpers
{
    public static string TempXlsxPath() =>
        Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslTest_{Guid.NewGuid():N}.xlsx");

    public static string TempXlsmPath() =>
        Path.Combine(Path.GetTempPath(), $"CsOpenXmlDslTest_{Guid.NewGuid():N}.xlsm");

    /// <summary>A real vbaProject.bin, shared with the F# test suite (extracted from a
    /// workbook actually saved by Excel) - see this project's own .csproj for how it's
    /// linked in rather than copy-pasted.</summary>
    public static byte[] SampleVbaProject() =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Assets", "sample.vbaProject.bin"));

    /// <summary>The canonical "1x1 transparent GIF" - the smallest possible valid image
    /// file, used ubiquitously as a web tracking pixel, so its bytes are about as
    /// well-known and trustworthy as test fixtures get. Same fixture the F# suite uses.</summary>
    public static byte[] OnePixelGif() =>
        Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7");

    /// <summary>Walks up from the test binary's output directory to find the repo root
    /// (marked by the solution file) - used to locate the wrapper's own .csproj for a
    /// <c>#:project</c> directive without a hard-coded absolute path.</summary>
    public static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "Kookerella.FsOpenXmlDsl.slnx")))
            dir = dir.Parent;

        return dir?.FullName ?? throw new InvalidOperationException($"Could not locate the repo root from {AppContext.BaseDirectory}");
    }

    public static void AssertSchemaValid(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }
}
