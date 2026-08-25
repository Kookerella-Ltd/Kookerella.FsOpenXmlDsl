namespace Kookerella.FsOpenXmlDsl

open System
open System.Xml
open System.Xml.Linq
open System.Xml.Schema

/// An XML surface over this DSL: `toWorkbook`/`ofWorkbook` translate between a plain
/// `Workbook` value and an `XElement` tree, so a caller who'd rather emit/consume XML than
/// write F# or C# (e.g. an XSLT pipeline producing a report) can build or inspect a workbook
/// without touching OOXML directly. Neither direction knows anything about OOXML - both go
/// through `Workbook`, the same value `Writer`/`Reader` do, so `Xml.toWorkbook >> Writer.write`
/// and `Reader.read >> Xml.ofWorkbook` are the actual save/load paths.
///
/// Convention: a discriminated union case becomes an XML element named after the case
/// (camelCased) when it carries data of its own (e.g. `CellValue.Formula` -> `<formula>`),
/// or an attribute *value* (also camelCased) when it's one of several parameterless
/// alternatives for a single property (e.g. `BorderLineStyle.Thin` -> `style="thin"`).
/// Optional fields are simply omitted, not written as empty/nil elements.
///
/// This module was built up feature-by-feature, the same order the C# wrapper was, and now
/// covers every feature the F# core models at the worksheet/workbook level: cell values,
/// styles, merged ranges, freeze panes, autofilter, column/row sizing, VBA (as opaque
/// base64 bytes), defined names, hyperlinks, comments, sheet/workbook protection, print
/// settings, images (also as opaque base64 bytes - this DSL doesn't decode/encode image
/// data any more than it does VBA), Excel Tables, sparklines, charts, pivot tables (note:
/// unlike everything else here, loading XML with a `<pivotTable>` doesn't re-run its
/// aggregation - the same as `Reader`, it's `Writer`/the `Sheet` builder path that actually
/// computes one; `toWorkbook` just carries the description through), conditional
/// formatting, and data validation.
module Xml =

    /// Loads the XSD schema this module's XML conforms to (`Xml.xsd`, embedded in this
    /// assembly rather than read from a repo-relative path, so it works for a consumer of
    /// the published package too). Compiled and ready to validate against, e.g.:
    /// `(Xml.ofWorkbook wb).Document.Validate(Xml.schemaSet (), fun _ e -> failwith e.Message)`.
    let schemaSet () : XmlSchemaSet =
        let assembly = Reflection.Assembly.GetExecutingAssembly()
        use stream = assembly.GetManifestResourceStream("Kookerella.FsOpenXmlDsl.Xml.xsd")

        if isNull stream then
            failwith "Embedded resource 'Kookerella.FsOpenXmlDsl.Xml.xsd' not found - check the EmbeddedResource/LogicalName wiring in the .fsproj."

        use reader = XmlReader.Create(stream)
        let set = XmlSchemaSet()
        set.Add(null, reader) |> ignore
        set.Compile()
        set

    // --- XML helpers -----------------------------------------------------------------
    //
    // `elem`/`elemA` take `obj list` (not `XObject list`) deliberately: F# list literals
    // don't implicitly upcast `XAttribute`/`XElement` items to a common `XObject` element
    // type the way C# collection initializers do, so every item is boxed explicitly at the
    // call site instead of relying on that covariance.

    let private xName (name: string) = XName.Get(name)

    let private childElem (name: string) (parent: XElement) : XElement option =
        parent.Element(xName name) |> Option.ofObj

    let private childElems (name: string) (parent: XElement) : XElement list =
        parent.Elements(xName name) |> List.ofSeq

    let private attrOpt (name: string) (e: XElement) : string option =
        e.Attribute(xName name) |> Option.ofObj |> Option.map (fun a -> a.Value)

    let private reqAttr (name: string) (e: XElement) : string =
        match attrOpt name e with
        | Some v -> v
        | None -> failwithf "<%s> is missing required attribute '%s'" e.Name.LocalName name

    let private boolAttr (name: string) (defaultValue: bool) (e: XElement) : bool =
        attrOpt name e |> Option.map bool.Parse |> Option.defaultValue defaultValue

    let private elem (name: string) (content: obj list) : XElement = XElement(xName name, List.toArray content)

    let private attr (name: string) (value: string) : XAttribute = XAttribute(xName name, value)

    /// F#'s own `string`/`ToString()` on a `bool` renders "True"/"False" - valid for
    /// `bool.Parse` (case-insensitive) but not for XSD's `xs:boolean` (lowercase only,
    /// per the spec), which `Xml.xsd` uses for every boolean attribute. Always go through
    /// this rather than `string`/`sprintf "%b"` for a bool that ends up as attribute text.
    let private ofBool (b: bool) : string = if b then "true" else "false"

    /// An attribute, boxed, only when the value is present - the common "optional attribute"
    /// shape used throughout this module.
    let private optAttr (name: string) (value: string option) : obj list =
        match value with
        | Some v -> [ box (attr name v) ]
        | None -> []

    let private boxElem (e: XElement option) : obj list = e |> Option.map box |> Option.toList

    // --- CellRef ---------------------------------------------------------------------

    let private ofCellRef (r: CellRef) : string = CellRef.toA1 r
    let private toCellRef (a1: string) : CellRef = CellRef.ofA1 a1

    // --- Color -------------------------------------------------------------------------

    let private toColor (e: XElement) : Color =
        match e.Name.LocalName with
        | "rgb" -> Rgb(byte (reqAttr "r" e), byte (reqAttr "g" e), byte (reqAttr "b" e))
        | "indexed" -> Indexed(int (reqAttr "value" e))
        | "theme" -> Theme(int (reqAttr "index" e), attrOpt "tint" e |> Option.map float)
        | other -> failwithf "Unknown <%s> - expected <rgb>, <indexed>, or <theme>" other

    let private ofColor (c: Color) : XElement =
        match c with
        | Rgb(r, g, b) -> elem "rgb" [ box (attr "r" (string r)); box (attr "g" (string g)); box (attr "b" (string b)) ]
        | Indexed i -> elem "indexed" [ box (attr "value" (string i)) ]
        | Theme(i, tint) -> elem "theme" (box (attr "index" (string i)) :: optAttr "tint" (tint |> Option.map string))

    // --- CellValue -----------------------------------------------------------------------

    let private toCellValue (parent: XElement) : CellValue =
        match parent.Elements() |> Seq.tryHead with
        | None -> Empty
        | Some e ->
            match e.Name.LocalName with
            | "text" -> Text(e.Value)
            | "number" -> Number(float e.Value)
            | "boolean" -> CellValue.Boolean(bool.Parse e.Value)
            | "date" -> Date(DateTime.Parse(e.Value, Globalization.CultureInfo.InvariantCulture))
            | "formula" -> Formula(e.Value, attrOpt "cachedValue" e |> Option.map float)
            | other -> failwithf "Unknown cell value element <%s>" other

    let private ofCellValue (v: CellValue) : XElement list =
        match v with
        | Empty -> []
        | Text s -> [ elem "text" [ box s ] ]
        | Number n -> [ elem "number" [ box (n.ToString(Globalization.CultureInfo.InvariantCulture)) ] ]
        | Boolean b -> [ elem "boolean" [ box (if b then "true" else "false") ] ]
        | Date d -> [ elem "date" [ box (d.ToString("o", Globalization.CultureInfo.InvariantCulture)) ] ]
        | Formula(expr, cached) -> [ elem "formula" (box expr :: optAttr "cachedValue" (cached |> Option.map string)) ]

    // --- Styles --------------------------------------------------------------------------

    let private toFontStyle (e: XElement) : FontStyle =
        { Name = attrOpt "name" e
          Size = attrOpt "size" e |> Option.map float
          Bold = boolAttr "bold" false e
          Italic = boolAttr "italic" false e
          Underline = boolAttr "underline" false e
          Strikethrough = boolAttr "strikethrough" false e
          Color = childElem "color" e |> Option.map (fun c -> c.Elements() |> Seq.head |> toColor) }

    let private ofFontStyle (f: FontStyle) : XElement =
        let attrs =
            optAttr "name" f.Name
            @ optAttr "size" (f.Size |> Option.map string)
            @ (if f.Bold then [ box (attr "bold" "true") ] else [])
            @ (if f.Italic then [ box (attr "italic" "true") ] else [])
            @ (if f.Underline then [ box (attr "underline" "true") ] else [])
            @ (if f.Strikethrough then [ box (attr "strikethrough" "true") ] else [])

        let colorChild = f.Color |> Option.map (fun c -> elem "color" [ box (ofColor c) ])
        elem "font" (attrs @ boxElem colorChild)

    let private toBorderLineStyle (s: string) : BorderLineStyle =
        match s with
        | "thin" -> Thin
        | "medium" -> Medium
        | "thick" -> Thick
        | "dashed" -> Dashed
        | "dotted" -> Dotted
        | "double" -> BorderLineStyle.Double
        | "hair" -> Hair
        | other -> Other other

    let private ofBorderLineStyle (s: BorderLineStyle) : string =
        match s with
        | Thin -> "thin"
        | Medium -> "medium"
        | Thick -> "thick"
        | Dashed -> "dashed"
        | Dotted -> "dotted"
        | Double -> "double"
        | Hair -> "hair"
        | Other s -> s

    let private toBorderSide (e: XElement) : BorderSide =
        { Style = toBorderLineStyle (reqAttr "style" e)
          Color = childElem "color" e |> Option.map (fun c -> c.Elements() |> Seq.head |> toColor) }

    let private ofBorderSide (s: BorderSide) (elementName: string) : XElement =
        let colorChild = s.Color |> Option.map (fun c -> elem "color" [ box (ofColor c) ])
        elem elementName (box (attr "style" (ofBorderLineStyle s.Style)) :: boxElem colorChild)

    let private toBorderStyle (e: XElement) : BorderStyle =
        { Left = childElem "left" e |> Option.map toBorderSide
          Right = childElem "right" e |> Option.map toBorderSide
          Top = childElem "top" e |> Option.map toBorderSide
          Bottom = childElem "bottom" e |> Option.map toBorderSide }

    let private ofBorderStyle (b: BorderStyle) : XElement =
        let sides =
            (b.Left |> Option.map (fun s -> ofBorderSide s "left") |> Option.toList)
            @ (b.Right |> Option.map (fun s -> ofBorderSide s "right") |> Option.toList)
            @ (b.Top |> Option.map (fun s -> ofBorderSide s "top") |> Option.toList)
            @ (b.Bottom |> Option.map (fun s -> ofBorderSide s "bottom") |> Option.toList)

        elem "border" (sides |> List.map box)

    let private toHorizontalAlignment (s: string) : HorizontalAlignment =
        match s with
        | "general" -> GeneralAlign
        | "left" -> AlignLeft
        | "center" -> AlignCenter
        | "right" -> AlignRight
        | "fill" -> AlignFill
        | "justify" -> AlignJustify
        | other -> failwithf "Unknown horizontal alignment '%s'" other

    let private ofHorizontalAlignment (a: HorizontalAlignment) : string =
        match a with
        | GeneralAlign -> "general"
        | AlignLeft -> "left"
        | AlignCenter -> "center"
        | AlignRight -> "right"
        | AlignFill -> "fill"
        | AlignJustify -> "justify"

    let private toVerticalAlignment (s: string) : VerticalAlignment =
        match s with
        | "top" -> AlignTop
        | "middle" -> AlignMiddle
        | "bottom" -> AlignBottom
        | other -> failwithf "Unknown vertical alignment '%s'" other

    let private ofVerticalAlignment (a: VerticalAlignment) : string =
        match a with
        | AlignTop -> "top"
        | AlignMiddle -> "middle"
        | AlignBottom -> "bottom"

    let private toAlignmentStyle (e: XElement) : AlignmentStyle =
        { Horizontal = attrOpt "horizontal" e |> Option.map toHorizontalAlignment
          Vertical = attrOpt "vertical" e |> Option.map toVerticalAlignment
          WrapText = boolAttr "wrapText" false e }

    let private ofAlignmentStyle (a: AlignmentStyle) : XElement =
        elem
            "alignment"
            (optAttr "horizontal" (a.Horizontal |> Option.map ofHorizontalAlignment)
             @ optAttr "vertical" (a.Vertical |> Option.map ofVerticalAlignment)
             @ (if a.WrapText then [ box (attr "wrapText" "true") ] else []))

    let private toNumberFormat (e: XElement) : NumberFormat =
        match attrOpt "custom" e with
        | Some code -> Custom code
        | None ->
            match reqAttr "kind" e with
            | "general" -> General
            | "integer" -> Integer
            | "twoDecimal" -> TwoDecimal
            | "percentage" -> Percentage
            | "currency" -> Currency
            | "shortDate" -> ShortDate
            | "dateAndTime" -> DateAndTime
            | other -> failwithf "Unknown number format kind '%s'" other

    let private ofNumberFormat (n: NumberFormat) : XElement =
        let kind =
            match n with
            | General -> Some "general"
            | Integer -> Some "integer"
            | TwoDecimal -> Some "twoDecimal"
            | Percentage -> Some "percentage"
            | Currency -> Some "currency"
            | ShortDate -> Some "shortDate"
            | DateAndTime -> Some "dateAndTime"
            | Custom _ -> None

        match n with
        | Custom code -> elem "numberFormat" [ box (attr "custom" code) ]
        | _ -> elem "numberFormat" (optAttr "kind" kind)

    let private toCellProtection (e: XElement) : CellProtection =
        { Locked = boolAttr "locked" true e
          Hidden = boolAttr "hidden" false e }

    let private ofCellProtection (p: CellProtection) : XElement =
        elem
            "protection"
            ((if not p.Locked then [ box (attr "locked" "false") ] else [])
             @ (if p.Hidden then [ box (attr "hidden" "true") ] else []))

    let private toCellStyle (e: XElement) : CellStyle =
        { Font = childElem "font" e |> Option.map toFontStyle
          Fill = childElem "fill" e |> Option.map (fun f -> { FillStyle.Color = f.Elements() |> Seq.head |> toColor })
          Border = childElem "border" e |> Option.map toBorderStyle
          NumberFormat = childElem "numberFormat" e |> Option.map toNumberFormat
          Alignment = childElem "alignment" e |> Option.map toAlignmentStyle
          Protection = childElem "protection" e |> Option.map toCellProtection }

    let private ofCellStyle (s: CellStyle) : XElement =
        let children =
            boxElem (s.Font |> Option.map ofFontStyle)
            @ boxElem (s.Fill |> Option.map (fun f -> elem "fill" [ box (ofColor f.Color) ]))
            @ boxElem (s.Border |> Option.map ofBorderStyle)
            @ boxElem (s.NumberFormat |> Option.map ofNumberFormat)
            @ boxElem (s.Alignment |> Option.map ofAlignmentStyle)
            @ boxElem (s.Protection |> Option.map ofCellProtection)

        elem "style" children

    // --- Cell / sheet-level structure ------------------------------------------------

    let private toCell (e: XElement) : Cell =
        { Ref = toCellRef (reqAttr "ref" e)
          Value = toCellValue e
          Style = childElem "style" e |> Option.map toCellStyle }

    let private ofCell (c: Cell) : XElement =
        let styleChild = c.Style |> Option.map ofCellStyle
        elem "cell" (box (attr "ref" (ofCellRef c.Ref)) :: (ofCellValue c.Value |> List.map box) @ boxElem styleChild)

    let private toMergedRange (e: XElement) : MergedRange =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e) }

    let private ofMergedRange (m: MergedRange) : XElement =
        elem "mergedRange" [ box (attr "topLeft" (ofCellRef m.TopLeft)); box (attr "bottomRight" (ofCellRef m.BottomRight)) ]

    let private toFreezePane (e: XElement) : FreezePane =
        { Rows = int (reqAttr "rows" e); Columns = int (reqAttr "columns" e) }

    let private ofFreezePane (f: FreezePane) : XElement =
        elem "freezePane" [ box (attr "rows" (string f.Rows)); box (attr "columns" (string f.Columns)) ]

    let private toAutoFilterRange (e: XElement) : AutoFilterRange =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e) }

    let private ofAutoFilterRange (a: AutoFilterRange) : XElement =
        elem "autoFilter" [ box (attr "topLeft" (ofCellRef a.TopLeft)); box (attr "bottomRight" (ofCellRef a.BottomRight)) ]

    let private toColumnProps (e: XElement) : int * ColumnProps =
        int (reqAttr "index" e), { Width = attrOpt "width" e |> Option.map float }

    let private ofColumnProps (index: int) (c: ColumnProps) : XElement =
        elem "columnProp" (box (attr "index" (string index)) :: optAttr "width" (c.Width |> Option.map string))

    let private toRowProps (e: XElement) : int * RowProps =
        int (reqAttr "index" e), { Height = attrOpt "height" e |> Option.map float }

    let private ofRowProps (index: int) (r: RowProps) : XElement =
        elem "rowProp" (box (attr "index" (string index)) :: optAttr "height" (r.Height |> Option.map string))

    // --- Comments ------------------------------------------------------------------------

    let private toCommentEntry (e: XElement) : CommentEntry =
        { Cell = toCellRef (reqAttr "cell" e)
          Author = attrOpt "author" e |> Option.defaultValue ""
          Text = e.Value }

    let private ofCommentEntry (c: CommentEntry) : XElement =
        elem "comment" (box (attr "cell" (ofCellRef c.Cell)) :: optAttr "author" (if c.Author = "" then None else Some c.Author) @ [ box c.Text ])

    // --- Hyperlinks ----------------------------------------------------------------------

    let private toHyperlinkTarget (e: XElement) : HyperlinkTarget =
        match e.Name.LocalName with
        | "externalHyperlink" -> ExternalHyperlink(e.Value)
        | "internalHyperlink" -> InternalHyperlink(e.Value)
        | other -> failwithf "Unknown <%s> - expected <externalHyperlink> or <internalHyperlink>" other

    let private ofHyperlinkTarget (t: HyperlinkTarget) : XElement =
        match t with
        | ExternalHyperlink url -> elem "externalHyperlink" [ box url ]
        | InternalHyperlink location -> elem "internalHyperlink" [ box location ]

    let private toHyperlinkEntry (e: XElement) : HyperlinkEntry =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e)
          Target = e.Elements() |> Seq.head |> toHyperlinkTarget
          Tooltip = attrOpt "tooltip" e
          Display = attrOpt "display" e }

    let private ofHyperlinkEntry (h: HyperlinkEntry) : XElement =
        elem
            "hyperlink"
            (box (attr "topLeft" (ofCellRef h.TopLeft))
             :: box (attr "bottomRight" (ofCellRef h.BottomRight))
             :: optAttr "tooltip" h.Tooltip
             @ optAttr "display" h.Display
             @ [ box (ofHyperlinkTarget h.Target) ])

    // --- PageSetup -----------------------------------------------------------------------

    let private toPageOrientation (s: string) : PageOrientation =
        match s with
        | "portrait" -> Portrait
        | "landscape" -> Landscape
        | other -> failwithf "Unknown page orientation '%s'" other

    let private ofPageOrientation (o: PageOrientation) : string =
        match o with
        | Portrait -> "portrait"
        | Landscape -> "landscape"

    let private toPaperSize (e: XElement) : PaperSize =
        match attrOpt "other" e with
        | Some code -> OtherPaperSize(int code)
        | None ->
            match reqAttr "kind" e with
            | "letter" -> Letter
            | "legal" -> Legal
            | "tabloid" -> Tabloid
            | "a3" -> A3
            | "a4" -> A4
            | other -> failwithf "Unknown paper size kind '%s'" other

    let private ofPaperSize (p: PaperSize) : XElement =
        match p with
        | Letter -> elem "paperSize" [ box (attr "kind" "letter") ]
        | Legal -> elem "paperSize" [ box (attr "kind" "legal") ]
        | Tabloid -> elem "paperSize" [ box (attr "kind" "tabloid") ]
        | A3 -> elem "paperSize" [ box (attr "kind" "a3") ]
        | A4 -> elem "paperSize" [ box (attr "kind" "a4") ]
        | OtherPaperSize code -> elem "paperSize" [ box (attr "other" (string code)) ]

    let private toPrintScaling (e: XElement) : PrintScaling =
        match attrOpt "percent" e with
        | Some p -> ScalePercent(int p)
        | None -> FitToPage(int (reqAttr "fitWidth" e), int (reqAttr "fitHeight" e))

    let private ofPrintScaling (s: PrintScaling) : XElement =
        match s with
        | ScalePercent p -> elem "scaling" [ box (attr "percent" (string p)) ]
        | FitToPage(w, h) -> elem "scaling" [ box (attr "fitWidth" (string w)); box (attr "fitHeight" (string h)) ]

    let private toPageMargins (e: XElement) : PageMargins =
        { Left = float (reqAttr "left" e)
          Right = float (reqAttr "right" e)
          Top = float (reqAttr "top" e)
          Bottom = float (reqAttr "bottom" e)
          Header = float (reqAttr "header" e)
          Footer = float (reqAttr "footer" e) }

    let private ofPageMargins (m: PageMargins) : XElement =
        elem
            "margins"
            [ box (attr "left" (string m.Left))
              box (attr "right" (string m.Right))
              box (attr "top" (string m.Top))
              box (attr "bottom" (string m.Bottom))
              box (attr "header" (string m.Header))
              box (attr "footer" (string m.Footer)) ]

    let private toPrintAreaRange (e: XElement) : CellRef * CellRef =
        toCellRef (reqAttr "topLeft" e), toCellRef (reqAttr "bottomRight" e)

    let private ofPrintAreaRange (topLeft: CellRef, bottomRight: CellRef) : XElement =
        elem "range" [ box (attr "topLeft" (ofCellRef topLeft)); box (attr "bottomRight" (ofCellRef bottomRight)) ]

    let private toPageSetup (e: XElement) : PageSetup =
        { Orientation = attrOpt "orientation" e |> Option.map toPageOrientation |> Option.defaultValue Portrait
          PaperSize = childElem "paperSize" e |> Option.map toPaperSize
          Scaling = childElem "scaling" e |> Option.map toPrintScaling
          Margins = childElem "margins" e |> Option.map toPageMargins |> Option.defaultValue PageMargins.Default
          PrintArea =
            childElem "printArea" e
            |> Option.map (childElems "range" >> List.map toPrintAreaRange)
            |> Option.defaultValue []
          Header = childElem "header" e |> Option.map (fun h -> h.Value)
          Footer = childElem "footer" e |> Option.map (fun h -> h.Value)
          EvenHeader = childElem "evenHeader" e |> Option.map (fun h -> h.Value)
          EvenFooter = childElem "evenFooter" e |> Option.map (fun h -> h.Value)
          FirstHeader = childElem "firstHeader" e |> Option.map (fun h -> h.Value)
          FirstFooter = childElem "firstFooter" e |> Option.map (fun h -> h.Value) }

    let private ofPageSetup (p: PageSetup) : XElement =
        let textChild name (v: string option) = v |> Option.map (fun s -> elem name [ box s ])

        let printAreaChild =
            if p.PrintArea.IsEmpty then
                None
            else
                Some(elem "printArea" (p.PrintArea |> List.map (ofPrintAreaRange >> box)))

        let children =
            boxElem (p.PaperSize |> Option.map ofPaperSize)
            @ boxElem (p.Scaling |> Option.map ofPrintScaling)
            @ [ box (ofPageMargins p.Margins) ]
            @ boxElem printAreaChild
            @ boxElem (textChild "header" p.Header)
            @ boxElem (textChild "footer" p.Footer)
            @ boxElem (textChild "evenHeader" p.EvenHeader)
            @ boxElem (textChild "evenFooter" p.EvenFooter)
            @ boxElem (textChild "firstHeader" p.FirstHeader)
            @ boxElem (textChild "firstFooter" p.FirstFooter)

        elem "pageSetup" (box (attr "orientation" (ofPageOrientation p.Orientation)) :: children)

    // --- Sparklines ----------------------------------------------------------------------

    let private toSparklineType (s: string) : SparklineType =
        match s with
        | "line" -> Line
        | "column" -> SparklineType.Column
        | "winLoss" -> WinLoss
        | other -> failwithf "Unknown sparkline type '%s'" other

    let private ofSparklineType (t: SparklineType) : string =
        match t with
        | Line -> "line"
        | SparklineType.Column -> "column"
        | WinLoss -> "winLoss"

    let private toSparklineStyle (e: XElement) : SparklineStyle =
        { Type = toSparklineType (reqAttr "type" e)
          Color = childElem "color" e |> Option.map (fun c -> c.Elements() |> Seq.head |> toColor)
          LineWeight = attrOpt "lineWeight" e |> Option.map float
          ShowMarkers = boolAttr "showMarkers" false e
          ShowHigh = boolAttr "showHigh" false e
          ShowLow = boolAttr "showLow" false e
          ShowFirst = boolAttr "showFirst" false e
          ShowLast = boolAttr "showLast" false e
          ShowNegative = boolAttr "showNegative" false e }

    let private ofSparklineStyle (s: SparklineStyle) : XElement =
        let colorChild = s.Color |> Option.map (fun c -> elem "color" [ box (ofColor c) ])

        elem
            "style"
            (box (attr "type" (ofSparklineType s.Type))
             :: optAttr "lineWeight" (s.LineWeight |> Option.map string)
             @ (if s.ShowMarkers then [ box (attr "showMarkers" "true") ] else [])
             @ (if s.ShowHigh then [ box (attr "showHigh" "true") ] else [])
             @ (if s.ShowLow then [ box (attr "showLow" "true") ] else [])
             @ (if s.ShowFirst then [ box (attr "showFirst" "true") ] else [])
             @ (if s.ShowLast then [ box (attr "showLast" "true") ] else [])
             @ (if s.ShowNegative then [ box (attr "showNegative" "true") ] else [])
             @ boxElem colorChild)

    let private toSparklineCell (e: XElement) : SparklineCell =
        { Cell = toCellRef (reqAttr "cell" e)
          DataTopLeft = toCellRef (reqAttr "dataTopLeft" e)
          DataBottomRight = toCellRef (reqAttr "dataBottomRight" e) }

    let private ofSparklineCell (s: SparklineCell) : XElement =
        elem
            "sparkline"
            [ box (attr "cell" (ofCellRef s.Cell))
              box (attr "dataTopLeft" (ofCellRef s.DataTopLeft))
              box (attr "dataBottomRight" (ofCellRef s.DataBottomRight)) ]

    let private toSparklineGroupEntry (e: XElement) : SparklineGroupEntry =
        { Style = childElem "style" e |> Option.map toSparklineStyle |> Option.defaultValue SparklineStyle.Default
          Sparklines =
            childElem "sparklines" e
            |> Option.map (childElems "sparkline" >> List.map toSparklineCell)
            |> Option.defaultValue [] }

    let private ofSparklineGroupEntry (g: SparklineGroupEntry) : XElement =
        elem
            "sparklineGroup"
            [ box (ofSparklineStyle g.Style)
              box (elem "sparklines" (g.Sparklines |> List.map (ofSparklineCell >> box))) ]

    // --- Charts --------------------------------------------------------------------------

    let private toChartType (s: string) : ChartType =
        match s with
        | "column" -> ChartColumn
        | "bar" -> ChartBar
        | "line" -> ChartLine
        | "pie" -> ChartPie
        | other -> failwithf "Unknown chart type '%s'" other

    let private ofChartType (t: ChartType) : string =
        match t with
        | ChartColumn -> "column"
        | ChartBar -> "bar"
        | ChartLine -> "line"
        | ChartPie -> "pie"

    let private toChartSeries (e: XElement) : ChartSeries =
        { Name = toCellRef (reqAttr "name" e)
          ValuesTopLeft = toCellRef (reqAttr "valuesTopLeft" e)
          ValuesBottomRight = toCellRef (reqAttr "valuesBottomRight" e) }

    let private ofChartSeries (s: ChartSeries) : XElement =
        elem
            "s"
            [ box (attr "name" (ofCellRef s.Name))
              box (attr "valuesTopLeft" (ofCellRef s.ValuesTopLeft))
              box (attr "valuesBottomRight" (ofCellRef s.ValuesBottomRight)) ]

    let private toChartEntry (e: XElement) : ChartEntry =
        let categoriesElem =
            match childElem "categories" e with
            | Some c -> c
            | None -> failwith "<chart> is missing required child <categories>"

        { Type = toChartType (reqAttr "type" e)
          Title = attrOpt "title" e
          CategoriesTopLeft = toCellRef (reqAttr "topLeft" categoriesElem)
          CategoriesBottomRight = toCellRef (reqAttr "bottomRight" categoriesElem)
          Series = childElem "series" e |> Option.map (childElems "s" >> List.map toChartSeries) |> Option.defaultValue []
          ShowLegend = boolAttr "showLegend" false e
          TopLeftAnchor = toCellRef (reqAttr "anchorTopLeft" e)
          BottomRightAnchor = toCellRef (reqAttr "anchorBottomRight" e) }

    let private ofChartEntry (c: ChartEntry) : XElement =
        elem
            "chart"
            (box (attr "type" (ofChartType c.Type))
             :: optAttr "title" c.Title
             @ (if c.ShowLegend then [ box (attr "showLegend" "true") ] else [])
             @ [ box (attr "anchorTopLeft" (ofCellRef c.TopLeftAnchor))
                 box (attr "anchorBottomRight" (ofCellRef c.BottomRightAnchor))
                 box (
                     elem
                         "categories"
                         [ box (attr "topLeft" (ofCellRef c.CategoriesTopLeft))
                           box (attr "bottomRight" (ofCellRef c.CategoriesBottomRight)) ]
                 )
                 box (elem "series" (c.Series |> List.map (ofChartSeries >> box))) ])

    // --- Pivot tables --------------------------------------------------------------------

    let private toPivotAggregation (s: string) : PivotAggregation =
        match s with
        | "sum" -> PivotSum
        | "count" -> PivotCount
        | "countNumbers" -> PivotCountNumbers
        | "average" -> PivotAverage
        | "min" -> PivotMin
        | "max" -> PivotMax
        | other -> failwithf "Unknown pivot aggregation '%s'" other

    let private ofPivotAggregation (a: PivotAggregation) : string =
        match a with
        | PivotSum -> "sum"
        | PivotCount -> "count"
        | PivotCountNumbers -> "countNumbers"
        | PivotAverage -> "average"
        | PivotMin -> "min"
        | PivotMax -> "max"

    let private toPivotTableEntry (e: XElement) : PivotTableEntry =
        { SourceSheet = attrOpt "sourceSheet" e
          SourceTopLeft = toCellRef (reqAttr "sourceTopLeft" e)
          SourceBottomRight = toCellRef (reqAttr "sourceBottomRight" e)
          RowField = reqAttr "rowField" e
          ColumnField = attrOpt "columnField" e
          ValueField = reqAttr "valueField" e
          Aggregation = toPivotAggregation (reqAttr "aggregation" e)
          ValueCaption = attrOpt "valueCaption" e
          TopLeftAnchor = toCellRef (reqAttr "anchorTopLeft" e) }

    let private ofPivotTableEntry (p: PivotTableEntry) : XElement =
        elem
            "pivotTable"
            (optAttr "sourceSheet" p.SourceSheet
             @ [ box (attr "sourceTopLeft" (ofCellRef p.SourceTopLeft))
                 box (attr "sourceBottomRight" (ofCellRef p.SourceBottomRight))
                 box (attr "rowField" p.RowField) ]
             @ optAttr "columnField" p.ColumnField
             @ [ box (attr "valueField" p.ValueField)
                 box (attr "aggregation" (ofPivotAggregation p.Aggregation)) ]
             @ optAttr "valueCaption" p.ValueCaption
             @ [ box (attr "anchorTopLeft" (ofCellRef p.TopLeftAnchor)) ])

    // --- ComparisonOperator (shared by conditional formatting and data validation) ------

    let private toComparisonOperator (s: string) : ComparisonOperator =
        match s with
        | "equal" -> Equal
        | "notEqual" -> NotEqual
        | "greaterThan" -> GreaterThan
        | "lessThan" -> LessThan
        | "greaterThanOrEqual" -> GreaterThanOrEqual
        | "lessThanOrEqual" -> LessThanOrEqual
        | "between" -> Between
        | "notBetween" -> NotBetween
        | other -> failwithf "Unknown comparison operator '%s'" other

    let private ofComparisonOperator (o: ComparisonOperator) : string =
        match o with
        | Equal -> "equal"
        | NotEqual -> "notEqual"
        | GreaterThan -> "greaterThan"
        | LessThan -> "lessThan"
        | GreaterThanOrEqual -> "greaterThanOrEqual"
        | LessThanOrEqual -> "lessThanOrEqual"
        | Between -> "between"
        | NotBetween -> "notBetween"

    // --- Conditional formatting ------------------------------------------------------

    let private toConditionalFormatRule (e: XElement) : ConditionalFormatRule =
        let ruleStyle () = childElem "style" e |> Option.map toCellStyle |> Option.defaultValue CellStyle.Default
        let color name = childElem name e |> Option.map (fun c -> c.Elements() |> Seq.head |> toColor) |> Option.get

        match e.Name.LocalName with
        | "cellValueRule" ->
            CellValueRule(toComparisonOperator (reqAttr "operator" e), reqAttr "formula1" e, attrOpt "formula2" e, ruleStyle ())
        | "formulaRule" -> FormulaRule(reqAttr "formula" e, ruleStyle ())
        | "colorScale2" -> ColorScale2(color "minColor", color "maxColor")
        | "colorScale3" -> ColorScale3(color "minColor", color "midColor", color "maxColor")
        | "dataBarRule" -> DataBarRule(color "color")
        | "duplicateValuesRule" -> DuplicateValuesRule(ruleStyle ())
        | "uniqueValuesRule" -> UniqueValuesRule(ruleStyle ())
        | other -> failwithf "Unknown conditional format rule <%s>" other

    let private ofConditionalFormatRule (r: ConditionalFormatRule) : XElement =
        let colorChild name (c: Color) = elem name [ box (ofColor c) ]

        match r with
        | CellValueRule(op, f1, f2, style) ->
            elem
                "cellValueRule"
                (box (attr "operator" (ofComparisonOperator op))
                 :: box (attr "formula1" f1)
                 :: optAttr "formula2" f2 @ [ box (ofCellStyle style) ])
        | FormulaRule(formula, style) -> elem "formulaRule" [ box (attr "formula" formula); box (ofCellStyle style) ]
        | ColorScale2(minColor, maxColor) -> elem "colorScale2" [ box (colorChild "minColor" minColor); box (colorChild "maxColor" maxColor) ]
        | ColorScale3(minColor, midColor, maxColor) ->
            elem
                "colorScale3"
                [ box (colorChild "minColor" minColor)
                  box (colorChild "midColor" midColor)
                  box (colorChild "maxColor" maxColor) ]
        | DataBarRule color -> elem "dataBarRule" [ box (colorChild "color" color) ]
        | DuplicateValuesRule style -> elem "duplicateValuesRule" [ box (ofCellStyle style) ]
        | UniqueValuesRule style -> elem "uniqueValuesRule" [ box (ofCellStyle style) ]

    let private toConditionalFormatEntry (e: XElement) : ConditionalFormatEntry =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e)
          Rule = e.Elements() |> Seq.head |> toConditionalFormatRule }

    let private ofConditionalFormatEntry (c: ConditionalFormatEntry) : XElement =
        elem
            "conditionalFormat"
            [ box (attr "topLeft" (ofCellRef c.TopLeft))
              box (attr "bottomRight" (ofCellRef c.BottomRight))
              box (ofConditionalFormatRule c.Rule) ]

    // --- Data validation -------------------------------------------------------------

    let private toErrorAlertStyle (s: string) : ErrorAlertStyle =
        match s with
        | "stop" -> Stop
        | "warning" -> Warning
        | "information" -> Information
        | other -> failwithf "Unknown error alert style '%s'" other

    let private ofErrorAlertStyle (s: ErrorAlertStyle) : string =
        match s with
        | Stop -> "stop"
        | Warning -> "warning"
        | Information -> "information"

    let private toValidationKind (e: XElement) : ValidationKind =
        match e.Name.LocalName with
        | "listValidation" -> ListValidation(childElems "item" e |> List.map (fun i -> i.Value))
        | "listFromRangeValidation" -> ListFromRangeValidation(toCellRef (reqAttr "topLeft" e), toCellRef (reqAttr "bottomRight" e))
        | "wholeNumberValidation" ->
            WholeNumberValidation(toComparisonOperator (reqAttr "operator" e), reqAttr "formula1" e, attrOpt "formula2" e)
        | "decimalValidation" -> DecimalValidation(toComparisonOperator (reqAttr "operator" e), reqAttr "formula1" e, attrOpt "formula2" e)
        | "textLengthValidation" ->
            TextLengthValidation(toComparisonOperator (reqAttr "operator" e), reqAttr "formula1" e, attrOpt "formula2" e)
        | "customValidation" -> CustomValidation(reqAttr "formula" e)
        | other -> failwithf "Unknown validation kind <%s>" other

    let private ofValidationKind (k: ValidationKind) : XElement =
        let comparisonKind name (op: ComparisonOperator) (f1: string) (f2: string option) =
            elem name (box (attr "operator" (ofComparisonOperator op)) :: box (attr "formula1" f1) :: optAttr "formula2" f2)

        match k with
        | ListValidation items -> elem "listValidation" (items |> List.map (fun i -> box (elem "item" [ box i ])))
        | ListFromRangeValidation(topLeft, bottomRight) ->
            elem "listFromRangeValidation" [ box (attr "topLeft" (ofCellRef topLeft)); box (attr "bottomRight" (ofCellRef bottomRight)) ]
        | WholeNumberValidation(op, f1, f2) -> comparisonKind "wholeNumberValidation" op f1 f2
        | DecimalValidation(op, f1, f2) -> comparisonKind "decimalValidation" op f1 f2
        | TextLengthValidation(op, f1, f2) -> comparisonKind "textLengthValidation" op f1 f2
        | CustomValidation formula -> elem "customValidation" [ box (attr "formula" formula) ]

    let private toValidationAlert (e: XElement) : ValidationAlert =
        { AllowBlank = boolAttr "allowBlank" true e
          ErrorStyle = attrOpt "errorStyle" e |> Option.map toErrorAlertStyle |> Option.defaultValue Stop
          ErrorTitle = attrOpt "errorTitle" e
          ErrorMessage = attrOpt "errorMessage" e
          InputTitle = attrOpt "inputTitle" e
          InputMessage = attrOpt "inputMessage" e }

    let private toDataValidationEntry (e: XElement) : DataValidationEntry =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e)
          Kind = e.Elements() |> Seq.head |> toValidationKind
          Alert = toValidationAlert e }

    let private ofDataValidationEntry (d: DataValidationEntry) : XElement =
        let alert = d.Alert

        elem
            "dataValidation"
            (box (attr "topLeft" (ofCellRef d.TopLeft))
             :: box (attr "bottomRight" (ofCellRef d.BottomRight))
             :: (if not alert.AllowBlank then [ box (attr "allowBlank" "false") ] else [])
             @ (if alert.ErrorStyle <> Stop then [ box (attr "errorStyle" (ofErrorAlertStyle alert.ErrorStyle)) ] else [])
             @ optAttr "errorTitle" alert.ErrorTitle
             @ optAttr "errorMessage" alert.ErrorMessage
             @ optAttr "inputTitle" alert.InputTitle
             @ optAttr "inputMessage" alert.InputMessage
             @ [ box (ofValidationKind d.Kind) ])

    // --- Images --------------------------------------------------------------------------

    let private toImageFormat (s: string) : ImageFormat =
        match s with
        | "png" -> Png
        | "jpeg" -> Jpeg
        | "gif" -> Gif
        | "bmp" -> Bmp
        | other -> failwithf "Unknown image format '%s'" other

    let private ofImageFormat (f: ImageFormat) : string =
        match f with
        | Png -> "png"
        | Jpeg -> "jpeg"
        | Gif -> "gif"
        | Bmp -> "bmp"

    let private toImageEntry (e: XElement) : ImageEntry =
        { Data = Convert.FromBase64String(e.Value)
          Format = toImageFormat (reqAttr "format" e)
          TopLeftAnchor = toCellRef (reqAttr "topLeft" e)
          BottomRightAnchor = toCellRef (reqAttr "bottomRight" e) }

    let private ofImageEntry (i: ImageEntry) : XElement =
        elem
            "image"
            [ box (attr "format" (ofImageFormat i.Format))
              box (attr "topLeft" (ofCellRef i.TopLeftAnchor))
              box (attr "bottomRight" (ofCellRef i.BottomRightAnchor))
              box (Convert.ToBase64String i.Data) ]

    // --- Tables --------------------------------------------------------------------------

    let private toTableColumn (e: XElement) : TableColumn =
        { Name = reqAttr "name" e
          CalculatedFormula = attrOpt "calculatedFormula" e }

    let private ofTableColumn (c: TableColumn) : XElement =
        elem "column" (box (attr "name" c.Name) :: optAttr "calculatedFormula" c.CalculatedFormula)

    let private toTableStyle (e: XElement) : TableStyle =
        { Name = attrOpt "name" e
          ShowFirstColumn = boolAttr "showFirstColumn" false e
          ShowLastColumn = boolAttr "showLastColumn" false e
          ShowRowStripes = boolAttr "showRowStripes" false e
          ShowColumnStripes = boolAttr "showColumnStripes" false e }

    let private ofTableStyle (s: TableStyle) : XElement =
        elem
            "style"
            (optAttr "name" s.Name
             @ (if s.ShowFirstColumn then [ box (attr "showFirstColumn" "true") ] else [])
             @ (if s.ShowLastColumn then [ box (attr "showLastColumn" "true") ] else [])
             @ (if s.ShowRowStripes then [ box (attr "showRowStripes" "true") ] else [])
             @ (if s.ShowColumnStripes then [ box (attr "showColumnStripes" "true") ] else []))

    let private toTableEntry (e: XElement) : TableEntry =
        { TopLeft = toCellRef (reqAttr "topLeft" e)
          BottomRight = toCellRef (reqAttr "bottomRight" e)
          Name = reqAttr "name" e
          Columns = childElem "columns" e |> Option.map (childElems "column" >> List.map toTableColumn) |> Option.defaultValue []
          Style = childElem "style" e |> Option.map toTableStyle |> Option.defaultValue TableStyle.Default }

    let private ofTableEntry (t: TableEntry) : XElement =
        elem
            "table"
            [ box (attr "topLeft" (ofCellRef t.TopLeft))
              box (attr "bottomRight" (ofCellRef t.BottomRight))
              box (attr "name" t.Name)
              box (elem "columns" (t.Columns |> List.map (ofTableColumn >> box)))
              box (ofTableStyle t.Style) ]

    // --- Protection ----------------------------------------------------------------------

    let private boolOptAttr (name: string) (e: XElement) : bool option = attrOpt name e |> Option.map bool.Parse

    let private optBoolAttr (name: string) (value: bool option) : obj list = optAttr name (value |> Option.map ofBool)

    let private toSheetProtection (e: XElement) : SheetProtection =
        { Password = attrOpt "password" e
          Sheet = boolAttr "sheet" true e
          Objects = boolOptAttr "objects" e
          Scenarios = boolOptAttr "scenarios" e
          FormatCells = boolOptAttr "formatCells" e
          FormatColumns = boolOptAttr "formatColumns" e
          FormatRows = boolOptAttr "formatRows" e
          InsertColumns = boolOptAttr "insertColumns" e
          InsertRows = boolOptAttr "insertRows" e
          InsertHyperlinks = boolOptAttr "insertHyperlinks" e
          DeleteColumns = boolOptAttr "deleteColumns" e
          DeleteRows = boolOptAttr "deleteRows" e
          SelectLockedCells = boolOptAttr "selectLockedCells" e
          Sort = boolOptAttr "sort" e
          AutoFilter = boolOptAttr "autoFilter" e
          PivotTables = boolOptAttr "pivotTables" e
          SelectUnlockedCells = boolOptAttr "selectUnlockedCells" e }

    let private ofSheetProtection (p: SheetProtection) : XElement =
        elem
            "protection"
            (optAttr "password" p.Password
             @ [ box (attr "sheet" (ofBool p.Sheet)) ]
             @ optBoolAttr "objects" p.Objects
             @ optBoolAttr "scenarios" p.Scenarios
             @ optBoolAttr "formatCells" p.FormatCells
             @ optBoolAttr "formatColumns" p.FormatColumns
             @ optBoolAttr "formatRows" p.FormatRows
             @ optBoolAttr "insertColumns" p.InsertColumns
             @ optBoolAttr "insertRows" p.InsertRows
             @ optBoolAttr "insertHyperlinks" p.InsertHyperlinks
             @ optBoolAttr "deleteColumns" p.DeleteColumns
             @ optBoolAttr "deleteRows" p.DeleteRows
             @ optBoolAttr "selectLockedCells" p.SelectLockedCells
             @ optBoolAttr "sort" p.Sort
             @ optBoolAttr "autoFilter" p.AutoFilter
             @ optBoolAttr "pivotTables" p.PivotTables
             @ optBoolAttr "selectUnlockedCells" p.SelectUnlockedCells)

    let private toWorkbookProtection (e: XElement) : WorkbookProtection =
        { Password = attrOpt "password" e
          LockStructure = boolOptAttr "lockStructure" e
          LockWindows = boolOptAttr "lockWindows" e }

    let private ofWorkbookProtection (p: WorkbookProtection) : XElement =
        elem
            "protection"
            (optAttr "password" p.Password
             @ optBoolAttr "lockStructure" p.LockStructure
             @ optBoolAttr "lockWindows" p.LockWindows)

    // --- Worksheet / Workbook ----------------------------------------------------------

    let private toWorksheet (e: XElement) : Worksheet =
        { Name = reqAttr "name" e
          Cells = childElem "cells" e |> Option.map (childElems "cell" >> List.map toCell) |> Option.defaultValue []
          ColumnProps =
            childElem "columnProps" e
            |> Option.map (childElems "columnProp" >> List.map toColumnProps >> Map.ofList)
            |> Option.defaultValue Map.empty
          RowProps =
            childElem "rowProps" e
            |> Option.map (childElems "rowProp" >> List.map toRowProps >> Map.ofList)
            |> Option.defaultValue Map.empty
          MergedRanges =
            childElem "mergedRanges" e
            |> Option.map (childElems "mergedRange" >> List.map toMergedRange)
            |> Option.defaultValue []
          FreezePane = childElem "freezePane" e |> Option.map toFreezePane
          AutoFilter = childElem "autoFilter" e |> Option.map toAutoFilterRange
          Protection = childElem "protection" e |> Option.map toSheetProtection
          ConditionalFormats =
            childElem "conditionalFormats" e
            |> Option.map (childElems "conditionalFormat" >> List.map toConditionalFormatEntry)
            |> Option.defaultValue []
          DataValidations =
            childElem "dataValidations" e
            |> Option.map (childElems "dataValidation" >> List.map toDataValidationEntry)
            |> Option.defaultValue []
          Hyperlinks =
            childElem "hyperlinks" e
            |> Option.map (childElems "hyperlink" >> List.map toHyperlinkEntry)
            |> Option.defaultValue []
          Comments =
            childElem "comments" e
            |> Option.map (childElems "comment" >> List.map toCommentEntry)
            |> Option.defaultValue []
          PageSetup = childElem "pageSetup" e |> Option.map toPageSetup
          Tables =
            childElem "tables" e
            |> Option.map (childElems "table" >> List.map toTableEntry)
            |> Option.defaultValue []
          SparklineGroups =
            childElem "sparklineGroups" e
            |> Option.map (childElems "sparklineGroup" >> List.map toSparklineGroupEntry)
            |> Option.defaultValue []
          Charts =
            childElem "charts" e
            |> Option.map (childElems "chart" >> List.map toChartEntry)
            |> Option.defaultValue []
          Images =
            childElem "images" e
            |> Option.map (childElems "image" >> List.map toImageEntry)
            |> Option.defaultValue []
          PivotTables =
            childElem "pivotTables" e
            |> Option.map (childElems "pivotTable" >> List.map toPivotTableEntry)
            |> Option.defaultValue [] }

    let private ofWorksheet (s: Worksheet) : XElement =
        let cellsChild = if s.Cells.IsEmpty then None else Some(elem "cells" (s.Cells |> List.map (ofCell >> box)))

        let mergedRangesChild =
            if s.MergedRanges.IsEmpty then
                None
            else
                Some(elem "mergedRanges" (s.MergedRanges |> List.map (ofMergedRange >> box)))

        let columnPropsChild =
            if s.ColumnProps.IsEmpty then
                None
            else
                Some(elem "columnProps" (s.ColumnProps |> Map.toList |> List.map (fun (i, c) -> box (ofColumnProps i c))))

        let rowPropsChild =
            if s.RowProps.IsEmpty then
                None
            else
                Some(elem "rowProps" (s.RowProps |> Map.toList |> List.map (fun (i, r) -> box (ofRowProps i r))))

        let hyperlinksChild =
            if s.Hyperlinks.IsEmpty then
                None
            else
                Some(elem "hyperlinks" (s.Hyperlinks |> List.map (ofHyperlinkEntry >> box)))

        let commentsChild =
            if s.Comments.IsEmpty then
                None
            else
                Some(elem "comments" (s.Comments |> List.map (ofCommentEntry >> box)))

        let sections =
            boxElem cellsChild
            @ boxElem mergedRangesChild
            @ boxElem (s.FreezePane |> Option.map ofFreezePane)
            @ boxElem (s.AutoFilter |> Option.map ofAutoFilterRange)
            @ boxElem columnPropsChild
            @ boxElem rowPropsChild
            @ boxElem hyperlinksChild
            @ boxElem commentsChild
            @ boxElem (s.PageSetup |> Option.map ofPageSetup)
            @ boxElem (s.Protection |> Option.map ofSheetProtection)
            @ boxElem (if s.Images.IsEmpty then None else Some(elem "images" (s.Images |> List.map (ofImageEntry >> box))))
            @ boxElem (if s.Tables.IsEmpty then None else Some(elem "tables" (s.Tables |> List.map (ofTableEntry >> box))))
            @ boxElem (
                if s.SparklineGroups.IsEmpty then
                    None
                else
                    Some(elem "sparklineGroups" (s.SparklineGroups |> List.map (ofSparklineGroupEntry >> box)))
            )
            @ boxElem (if s.Charts.IsEmpty then None else Some(elem "charts" (s.Charts |> List.map (ofChartEntry >> box))))
            @ boxElem (
                if s.PivotTables.IsEmpty then
                    None
                else
                    Some(elem "pivotTables" (s.PivotTables |> List.map (ofPivotTableEntry >> box)))
            )
            @ boxElem (
                if s.ConditionalFormats.IsEmpty then
                    None
                else
                    Some(elem "conditionalFormats" (s.ConditionalFormats |> List.map (ofConditionalFormatEntry >> box)))
            )
            @ boxElem (
                if s.DataValidations.IsEmpty then
                    None
                else
                    Some(elem "dataValidations" (s.DataValidations |> List.map (ofDataValidationEntry >> box)))
            )

        elem "sheet" (box (attr "name" s.Name) :: sections)

    // --- DefinedNames --------------------------------------------------------------------

    let private toDefinedNameScope (e: XElement) : DefinedNameScope =
        match e.Name.LocalName with
        | "workbookScope" -> WorkbookScope
        | "sheetScope" -> SheetScope(reqAttr "sheetName" e)
        | other -> failwithf "Unknown <%s> - expected <workbookScope> or <sheetScope>" other

    let private ofDefinedNameScope (s: DefinedNameScope) : XElement =
        match s with
        | WorkbookScope -> elem "workbookScope" []
        | SheetScope sheetName -> elem "sheetScope" [ box (attr "sheetName" sheetName) ]

    let private toDefinedNameEntry (e: XElement) : DefinedNameEntry =
        { Name = reqAttr "name" e
          Formula = reqAttr "formula" e
          Scope = e.Elements() |> Seq.head |> toDefinedNameScope
          Hidden = boolAttr "hidden" false e }

    let private ofDefinedNameEntry (d: DefinedNameEntry) : XElement =
        elem
            "definedName"
            (box (attr "name" d.Name)
             :: box (attr "formula" d.Formula)
             :: box (ofDefinedNameScope d.Scope)
             :: (if d.Hidden then [ box (attr "hidden" "true") ] else []))

    /// Reads a `Workbook` from an XML tree. `root` should be a `<workbook>` element - see
    /// this module's own doc comment for the schema `ofWorkbook` produces (the two are each
    /// other's exact inverse for anything this pass covers).
    let toWorkbook (root: XElement) : Workbook =
        { Sheets = childElem "sheets" root |> Option.map (childElems "sheet" >> List.map toWorksheet) |> Option.defaultValue []
          DefinedNames =
            childElem "definedNames" root
            |> Option.map (childElems "definedName" >> List.map toDefinedNameEntry)
            |> Option.defaultValue []
          Protection = childElem "protection" root |> Option.map toWorkbookProtection
          VbaProject = childElem "vbaProject" root |> Option.map (fun e -> Convert.FromBase64String(e.Value)) }

    /// Renders a `Workbook` as an XML tree rooted at `<workbook>`.
    let ofWorkbook (wb: Workbook) : XElement =
        let sheetsChild = elem "sheets" (wb.Sheets |> List.map (ofWorksheet >> box))

        let definedNamesChild =
            if wb.DefinedNames.IsEmpty then
                None
            else
                Some(elem "definedNames" (wb.DefinedNames |> List.map (ofDefinedNameEntry >> box)))

        let protectionChild = wb.Protection |> Option.map ofWorkbookProtection
        let vbaChild = wb.VbaProject |> Option.map (fun bytes -> elem "vbaProject" [ box (Convert.ToBase64String bytes) ])

        elem "workbook" (box sheetsChild :: boxElem definedNamesChild @ boxElem protectionChild @ boxElem vbaChild)
