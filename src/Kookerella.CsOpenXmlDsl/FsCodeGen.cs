using Microsoft.FSharp.Collections;
using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Renders a <see cref="Workbook"/> back out as a self-contained F# script that rebuilds an
/// equivalent file when run via <c>dotnet fsi</c> - the F#-side equivalent of
/// <see cref="CsCodeGen.Generate"/>, for a caller who built a workbook through this C#
/// wrapper but wants F# output rather than C#. A thin wrapper over the F# core's own
/// <c>Workbook.generateScript</c> (going through <see cref="WorkbookConverter"/> first), not
/// a separate reimplementation the way <see cref="CsCodeGen"/> is - F# source is F# source
/// regardless of which language built the source <see cref="Workbook"/>, so there's no
/// C#-specific rendering work to do here.
/// </summary>
public static class FsCodeGen
{
    public static string Generate(IReadOnlyList<string> referenceLines, string outputFileName, Workbook workbook) =>
        Fs.Workbook.generateScript(
            ListModule.OfSeq(referenceLines),
            outputFileName,
            WorkbookConverter.ToFSharp(workbook));
}
