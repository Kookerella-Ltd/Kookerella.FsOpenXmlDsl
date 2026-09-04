using System.Xml.Linq;
using Fs = Kookerella.FsOpenXmlDsl;

namespace Kookerella.CsOpenXmlDsl;

/// <summary>
/// Converts a <see cref="Workbook"/> to/from XML matching the F# core's own embedded schema
/// (<c>Xml.xsd</c>) - a thin wrapper over <c>Kookerella.FsOpenXmlDsl.Xml.ofWorkbook</c>/
/// <c>toWorkbook</c>, going through <see cref="WorkbookConverter"/> so this assembly's own
/// types never need to touch the F# core's directly.
/// </summary>
public static class WorkbookXml
{
    public static XElement ToXml(Workbook workbook) =>
        Fs.Xml.ofWorkbook(WorkbookConverter.ToFSharp(workbook));

    public static Workbook FromXml(XElement element) =>
        WorkbookConverter.FromFSharp(Fs.Xml.toWorkbook(element));
}
