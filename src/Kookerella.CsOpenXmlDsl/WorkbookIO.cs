using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// The one place this wrapper does anything side-effecting - every other type in this
/// assembly is a pure, immutable value with no I/O methods of its own (see <see
/// cref="Workbook"/>'s own doc comment). Mirrors the F# core's own separation between the
/// <c>Workbook</c> data type and its <c>Workbook.save</c>/<c>load</c> module functions.
/// </summary>
public static class WorkbookIO
{
    public static void Save(Workbook workbook, string path) =>
        Fs.Workbook.save(path, WorkbookConverter.ToFSharp(workbook));

    public static void Save(Workbook workbook, Stream stream) =>
        Fs.Workbook.saveToStream(stream, WorkbookConverter.ToFSharp(workbook));

    public static Workbook Load(string path) =>
        WorkbookConverter.FromFSharp(Fs.Workbook.load(path));

    public static Workbook Load(Stream stream) =>
        WorkbookConverter.FromFSharp(Fs.Workbook.loadFromStream(stream));
}
