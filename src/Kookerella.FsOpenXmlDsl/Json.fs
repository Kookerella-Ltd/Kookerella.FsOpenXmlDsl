namespace Kookerella.FsOpenXmlDsl

open System
open System.Text.Json.Nodes

/// A JSON surface over this DSL, alongside `Xml.fs`: `toWorkbook`/`ofWorkbook` translate
/// between a plain `Workbook` value and a `System.Text.Json.Nodes.JsonObject` tree. Same
/// motivation as `Xml.fs` (a caller who'd rather generate/consume data than write code),
/// different wire format for callers whose tooling speaks JSON rather than XML.
///
/// Convention, translated from `Xml.fs`'s own into JSON's shape: a discriminated union case
/// becomes a single-key JSON object named after the case (camelCased) when it carries data
/// of its own (e.g. `CellValue.Formula` -> `{"formula": {...}}`), or a bare JSON string
/// value (also camelCased) when it's one of several parameterless alternatives (e.g.
/// `BorderLineStyle.Thin` -> `"style": "thin"`). Optional fields are simply omitted, not
/// written as `null`. Unlike XML, JSON has no schema built into .NET the way
/// `System.Xml.Schema` is - `Json.schema.json` (hand-written, matching this module's actual
/// shape) is validated only from the test suite (via a test-only JsonSchema.Net dependency),
/// not exposed as a public API the way `Xml.schemaSet()` is.
///
/// Built up feature-by-feature, the same order `Xml.fs` was: this pass covers the
/// foundation (cell values, styles, merged ranges, freeze panes, autofilter, column/row
/// sizing), VBA, defined names, comments, hyperlinks, sheet/workbook protection, page
/// setup, images, tables, sparklines, charts, pivot tables, conditional formatting, and
/// data validation - the same feature set `Xml.fs` covers.
module Json =

    // --- JSON helpers ------------------------------------------------------------------

    let private tryGet (key: string) (o: JsonObject) : JsonNode option =
        let mutable node: JsonNode = null

        if o.TryGetPropertyValue(key, &node) && not (isNull node) then
            Some node
        else
            None

    let private childObj (key: string) (o: JsonObject) : JsonObject option = tryGet key o |> Option.map (fun n -> n.AsObject())

    let private childArr (key: string) (o: JsonObject) : JsonArray option = tryGet key o |> Option.map (fun n -> n.AsArray())

    let private reqStr (key: string) (o: JsonObject) : string =
        match tryGet key o with
        | Some n -> n.GetValue<string>()
        | None -> failwithf "<%s> is missing required property '%s'" (o.ToJsonString()) key

    let private optStr (key: string) (o: JsonObject) : string option = tryGet key o |> Option.map (fun n -> n.GetValue<string>())

    let private reqNum (key: string) (o: JsonObject) : float =
        match tryGet key o with
        | Some n -> n.GetValue<float>()
        | None -> failwithf "<%s> is missing required property '%s'" (o.ToJsonString()) key

    let private optNum (key: string) (o: JsonObject) : float option = tryGet key o |> Option.map (fun n -> n.GetValue<float>())

    let private reqInt (key: string) (o: JsonObject) : int =
        match tryGet key o with
        | Some n -> n.GetValue<int>()
        | None -> failwithf "<%s> is missing required property '%s'" (o.ToJsonString()) key

    let private optBool (key: string) (defaultValue: bool) (o: JsonObject) : bool =
        tryGet key o |> Option.map (fun n -> n.GetValue<bool>()) |> Option.defaultValue defaultValue

    let private optBoolOption (key: string) (o: JsonObject) : bool option = tryGet key o |> Option.map (fun n -> n.GetValue<bool>())

    /// Builds a `JsonObject` from a list of (key, value option) pairs, omitting any `None`
    /// - the JSON-side equivalent of `Xml.fs`'s `optAttr`/`boxElem` combination.
    let private obj (pairs: (string * JsonNode option) list) : JsonObject =
        let o = JsonObject()

        for key, value in pairs do
            match value with
            | Some node -> o.[key] <- node
            | None -> ()

        o

    let private req (key: string) (node: JsonNode) : string * JsonNode option = key, Some node
    let private optNode (key: string) (value: JsonNode option) : string * JsonNode option = key, value
    let private str (value: string) : JsonNode = JsonValue.Create(value)
    let private num (value: float) : JsonNode = JsonValue.Create(value)
    let private jint (value: int) : JsonNode = JsonValue.Create(value)
    let private jbool (value: bool) : JsonNode = JsonValue.Create(value)
    let private jarr (items: JsonNode list) : JsonNode = JsonArray(items |> List.toArray) :> JsonNode

    // --- CellRef -------------------------------------------------------------------------

    let private ofCellRef (r: CellRef) : string = CellRef.toA1 r
    let private toCellRef (a1: string) : CellRef = CellRef.ofA1 a1

    // --- Color -------------------------------------------------------------------------

    let private toColor (o: JsonObject) : Color =
        match tryGet "rgb" o with
        | Some rgbNode ->
            let rgb = rgbNode.AsObject()
            Rgb(byte (reqInt "r" rgb), byte (reqInt "g" rgb), byte (reqInt "b" rgb))
        | None ->
            match tryGet "indexed" o with
            | Some n -> Indexed(n.GetValue<int>())
            | None ->
                match tryGet "theme" o with
                | Some themeNode ->
                    let theme = themeNode.AsObject()
                    Theme(reqInt "index" theme, optNum "tint" theme)
                | None -> failwithf "Expected one of 'rgb', 'indexed', 'theme' in %s" (o.ToJsonString())

    let private ofColor (c: Color) : JsonNode =
        match c with
        | Rgb(r, g, b) -> obj [ req "rgb" (obj [ req "r" (jint (int r)); req "g" (jint (int g)); req "b" (jint (int b)) ]) ] :> JsonNode
        | Indexed i -> obj [ req "indexed" (jint i) ] :> JsonNode
        | Theme(i, tint) -> obj [ req "theme" (obj [ req "index" (jint i); optNode "tint" (tint |> Option.map num) ]) ] :> JsonNode

    // --- CellValue -----------------------------------------------------------------------

    let private toCellValue (o: JsonObject) : CellValue =
        match tryGet "text" o with
        | Some n -> Text(n.GetValue<string>())
        | None ->
            match tryGet "number" o with
            | Some n -> Number(n.GetValue<float>())
            | None ->
                match tryGet "boolean" o with
                | Some n -> CellValue.Boolean(n.GetValue<bool>())
                | None ->
                    match tryGet "date" o with
                    | Some n -> Date(DateTime.Parse(n.GetValue<string>(), Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.RoundtripKind))
                    | None ->
                        match tryGet "formula" o with
                        | Some n ->
                            let f = n.AsObject()
                            Formula(reqStr "expression" f, optNum "cachedValue" f)
                        | None -> Empty

    let private ofCellValueFields (v: CellValue) : (string * JsonNode option) list =
        match v with
        | Empty -> []
        | Text s -> [ req "text" (str s) ]
        | Number n -> [ req "number" (num n) ]
        | CellValue.Boolean b -> [ req "boolean" (jbool b) ]
        | Date d -> [ req "date" (str (d.ToString("o", Globalization.CultureInfo.InvariantCulture))) ]
        | Formula(expr, cached) -> [ req "formula" (obj [ req "expression" (str expr); optNode "cachedValue" (cached |> Option.map num) ]) ]

    // --- Styles --------------------------------------------------------------------------

    let private toFontStyle (o: JsonObject) : FontStyle =
        { Name = optStr "name" o
          Size = optNum "size" o
          Bold = optBool "bold" false o
          Italic = optBool "italic" false o
          Underline = optBool "underline" false o
          Strikethrough = optBool "strikethrough" false o
          Color = childObj "color" o |> Option.map toColor }

    let private ofFontStyle (f: FontStyle) : JsonNode =
        obj
            [ optNode "name" (f.Name |> Option.map str)
              optNode "size" (f.Size |> Option.map num)
              (if f.Bold then req "bold" (jbool true) else "bold", None)
              (if f.Italic then req "italic" (jbool true) else "italic", None)
              (if f.Underline then req "underline" (jbool true) else "underline", None)
              (if f.Strikethrough then req "strikethrough" (jbool true) else "strikethrough", None)
              optNode "color" (f.Color |> Option.map ofColor) ]
        :> JsonNode

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
        | BorderLineStyle.Double -> "double"
        | Hair -> "hair"
        | Other s -> s

    let private toBorderSide (o: JsonObject) : BorderSide =
        { Style = toBorderLineStyle (reqStr "style" o)
          Color = childObj "color" o |> Option.map toColor }

    let private ofBorderSide (s: BorderSide) : JsonNode =
        obj [ req "style" (str (ofBorderLineStyle s.Style)); optNode "color" (s.Color |> Option.map ofColor) ] :> JsonNode

    let private toBorderStyle (o: JsonObject) : BorderStyle =
        { Left = childObj "left" o |> Option.map toBorderSide
          Right = childObj "right" o |> Option.map toBorderSide
          Top = childObj "top" o |> Option.map toBorderSide
          Bottom = childObj "bottom" o |> Option.map toBorderSide }

    let private ofBorderStyle (b: BorderStyle) : JsonNode =
        obj
            [ optNode "left" (b.Left |> Option.map ofBorderSide)
              optNode "right" (b.Right |> Option.map ofBorderSide)
              optNode "top" (b.Top |> Option.map ofBorderSide)
              optNode "bottom" (b.Bottom |> Option.map ofBorderSide) ]
        :> JsonNode

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

    let private toAlignmentStyle (o: JsonObject) : AlignmentStyle =
        { Horizontal = optStr "horizontal" o |> Option.map toHorizontalAlignment
          Vertical = optStr "vertical" o |> Option.map toVerticalAlignment
          WrapText = optBool "wrapText" false o }

    let private ofAlignmentStyle (a: AlignmentStyle) : JsonNode =
        obj
            [ optNode "horizontal" (a.Horizontal |> Option.map (ofHorizontalAlignment >> str))
              optNode "vertical" (a.Vertical |> Option.map (ofVerticalAlignment >> str))
              (if a.WrapText then req "wrapText" (jbool true) else "wrapText", None) ]
        :> JsonNode

    let private toNumberFormat (node: JsonNode) : NumberFormat =
        match node with
        | :? JsonValue as v when v.TryGetValue<string>() |> fst ->
            match v.GetValue<string>() with
            | "general" -> General
            | "integer" -> Integer
            | "twoDecimal" -> TwoDecimal
            | "percentage" -> Percentage
            | "currency" -> Currency
            | "shortDate" -> ShortDate
            | "dateAndTime" -> DateAndTime
            | other -> failwithf "Unknown number format kind '%s'" other
        | _ -> Custom(reqStr "custom" (node.AsObject()))

    let private ofNumberFormat (n: NumberFormat) : JsonNode =
        match n with
        | General -> str "general"
        | Integer -> str "integer"
        | TwoDecimal -> str "twoDecimal"
        | Percentage -> str "percentage"
        | Currency -> str "currency"
        | ShortDate -> str "shortDate"
        | DateAndTime -> str "dateAndTime"
        | Custom code -> obj [ req "custom" (str code) ] :> JsonNode

    let private toCellProtection (o: JsonObject) : CellProtection =
        { Locked = optBool "locked" true o
          Hidden = optBool "hidden" false o }

    let private ofCellProtection (p: CellProtection) : JsonNode =
        obj
            [ (if not p.Locked then req "locked" (jbool false) else "locked", None)
              (if p.Hidden then req "hidden" (jbool true) else "hidden", None) ]
        :> JsonNode

    let private toCellStyle (o: JsonObject) : CellStyle =
        { Font = childObj "font" o |> Option.map toFontStyle
          Fill = childObj "fill" o |> Option.map (fun f -> { FillStyle.Color = toColor ((childObj "color" f).Value) })
          Border = childObj "border" o |> Option.map toBorderStyle
          NumberFormat = tryGet "numberFormat" o |> Option.map toNumberFormat
          Alignment = childObj "alignment" o |> Option.map toAlignmentStyle
          Protection = childObj "protection" o |> Option.map toCellProtection }

    let private ofCellStyle (s: CellStyle) : JsonNode =
        obj
            [ optNode "font" (s.Font |> Option.map ofFontStyle)
              optNode "fill" (s.Fill |> Option.map (fun f -> obj [ req "color" (ofColor f.Color) ] :> JsonNode))
              optNode "border" (s.Border |> Option.map ofBorderStyle)
              optNode "numberFormat" (s.NumberFormat |> Option.map ofNumberFormat)
              optNode "alignment" (s.Alignment |> Option.map ofAlignmentStyle)
              optNode "protection" (s.Protection |> Option.map ofCellProtection) ]
        :> JsonNode

    // --- Cell / sheet-level structure ------------------------------------------------

    let private toCell (o: JsonObject) : Cell =
        { Ref = toCellRef (reqStr "ref" o)
          Value = toCellValue o
          Style = childObj "style" o |> Option.map toCellStyle }

    let private ofCell (c: Cell) : JsonNode =
        obj (req "ref" (str (ofCellRef c.Ref)) :: ofCellValueFields c.Value @ [ optNode "style" (c.Style |> Option.map ofCellStyle) ]) :> JsonNode

    let private toMergedRange (o: JsonObject) : MergedRange =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o) }

    let private ofMergedRange (m: MergedRange) : JsonNode =
        obj [ req "topLeft" (str (ofCellRef m.TopLeft)); req "bottomRight" (str (ofCellRef m.BottomRight)) ] :> JsonNode

    let private toFreezePane (o: JsonObject) : FreezePane =
        { Rows = reqInt "rows" o; Columns = reqInt "columns" o }

    let private ofFreezePane (f: FreezePane) : JsonNode =
        obj [ req "rows" (jint f.Rows); req "columns" (jint f.Columns) ] :> JsonNode

    let private toAutoFilterRange (o: JsonObject) : AutoFilterRange =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o) }

    let private ofAutoFilterRange (a: AutoFilterRange) : JsonNode =
        obj [ req "topLeft" (str (ofCellRef a.TopLeft)); req "bottomRight" (str (ofCellRef a.BottomRight)) ] :> JsonNode

    let private toColumnProps (o: JsonObject) : int * ColumnProps = reqInt "index" o, { Width = optNum "width" o }

    let private ofColumnProps (index: int) (c: ColumnProps) : JsonNode =
        obj [ req "index" (jint index); optNode "width" (c.Width |> Option.map num) ] :> JsonNode

    let private toRowProps (o: JsonObject) : int * RowProps = reqInt "index" o, { Height = optNum "height" o }

    let private ofRowProps (index: int) (r: RowProps) : JsonNode =
        obj [ req "index" (jint index); optNode "height" (r.Height |> Option.map num) ] :> JsonNode

    // --- Comments ------------------------------------------------------------------------

    let private toCommentEntry (o: JsonObject) : CommentEntry =
        { Cell = toCellRef (reqStr "cell" o)
          Author = optStr "author" o |> Option.defaultValue ""
          Text = reqStr "text" o }

    let private ofCommentEntry (c: CommentEntry) : JsonNode =
        obj
            [ req "cell" (str (ofCellRef c.Cell))
              optNode "author" (if c.Author = "" then None else Some(str c.Author))
              req "text" (str c.Text) ]
        :> JsonNode

    // --- Hyperlinks ----------------------------------------------------------------------

    let private toHyperlinkTarget (o: JsonObject) : HyperlinkTarget =
        match tryGet "externalHyperlink" o with
        | Some n -> ExternalHyperlink(n.GetValue<string>())
        | None ->
            match tryGet "internalHyperlink" o with
            | Some n -> InternalHyperlink(n.GetValue<string>())
            | None -> failwithf "Expected 'externalHyperlink' or 'internalHyperlink' in %s" (o.ToJsonString())

    let private ofHyperlinkTarget (t: HyperlinkTarget) : JsonNode =
        match t with
        | ExternalHyperlink url -> obj [ req "externalHyperlink" (str url) ] :> JsonNode
        | InternalHyperlink location -> obj [ req "internalHyperlink" (str location) ] :> JsonNode

    let private toHyperlinkEntry (o: JsonObject) : HyperlinkEntry =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o)
          Target = toHyperlinkTarget ((childObj "target" o).Value)
          Tooltip = optStr "tooltip" o
          Display = optStr "display" o }

    let private ofHyperlinkEntry (h: HyperlinkEntry) : JsonNode =
        obj
            [ req "topLeft" (str (ofCellRef h.TopLeft))
              req "bottomRight" (str (ofCellRef h.BottomRight))
              req "target" (ofHyperlinkTarget h.Target)
              optNode "tooltip" (h.Tooltip |> Option.map str)
              optNode "display" (h.Display |> Option.map str) ]
        :> JsonNode

    // --- Protection ----------------------------------------------------------------------

    let private toSheetProtection (o: JsonObject) : SheetProtection =
        { Password = optStr "password" o
          Sheet = optBool "sheet" true o
          Objects = optBoolOption "objects" o
          Scenarios = optBoolOption "scenarios" o
          FormatCells = optBoolOption "formatCells" o
          FormatColumns = optBoolOption "formatColumns" o
          FormatRows = optBoolOption "formatRows" o
          InsertColumns = optBoolOption "insertColumns" o
          InsertRows = optBoolOption "insertRows" o
          InsertHyperlinks = optBoolOption "insertHyperlinks" o
          DeleteColumns = optBoolOption "deleteColumns" o
          DeleteRows = optBoolOption "deleteRows" o
          SelectLockedCells = optBoolOption "selectLockedCells" o
          Sort = optBoolOption "sort" o
          AutoFilter = optBoolOption "autoFilter" o
          PivotTables = optBoolOption "pivotTables" o
          SelectUnlockedCells = optBoolOption "selectUnlockedCells" o }

    let private ofSheetProtection (p: SheetProtection) : JsonNode =
        obj
            [ optNode "password" (p.Password |> Option.map str)
              req "sheet" (jbool p.Sheet)
              optNode "objects" (p.Objects |> Option.map jbool)
              optNode "scenarios" (p.Scenarios |> Option.map jbool)
              optNode "formatCells" (p.FormatCells |> Option.map jbool)
              optNode "formatColumns" (p.FormatColumns |> Option.map jbool)
              optNode "formatRows" (p.FormatRows |> Option.map jbool)
              optNode "insertColumns" (p.InsertColumns |> Option.map jbool)
              optNode "insertRows" (p.InsertRows |> Option.map jbool)
              optNode "insertHyperlinks" (p.InsertHyperlinks |> Option.map jbool)
              optNode "deleteColumns" (p.DeleteColumns |> Option.map jbool)
              optNode "deleteRows" (p.DeleteRows |> Option.map jbool)
              optNode "selectLockedCells" (p.SelectLockedCells |> Option.map jbool)
              optNode "sort" (p.Sort |> Option.map jbool)
              optNode "autoFilter" (p.AutoFilter |> Option.map jbool)
              optNode "pivotTables" (p.PivotTables |> Option.map jbool)
              optNode "selectUnlockedCells" (p.SelectUnlockedCells |> Option.map jbool) ]
        :> JsonNode

    let private toWorkbookProtection (o: JsonObject) : WorkbookProtection =
        { Password = optStr "password" o
          LockStructure = optBoolOption "lockStructure" o
          LockWindows = optBoolOption "lockWindows" o }

    let private ofWorkbookProtection (p: WorkbookProtection) : JsonNode =
        obj
            [ optNode "password" (p.Password |> Option.map str)
              optNode "lockStructure" (p.LockStructure |> Option.map jbool)
              optNode "lockWindows" (p.LockWindows |> Option.map jbool) ]
        :> JsonNode

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

    let private toPaperSize (node: JsonNode) : PaperSize =
        match node with
        | :? JsonValue as v when v.TryGetValue<string>() |> fst ->
            match v.GetValue<string>() with
            | "letter" -> Letter
            | "legal" -> Legal
            | "tabloid" -> Tabloid
            | "a3" -> A3
            | "a4" -> A4
            | other -> failwithf "Unknown paper size kind '%s'" other
        | _ -> OtherPaperSize(reqInt "other" (node.AsObject()))

    let private ofPaperSize (p: PaperSize) : JsonNode =
        match p with
        | Letter -> str "letter"
        | Legal -> str "legal"
        | Tabloid -> str "tabloid"
        | A3 -> str "a3"
        | A4 -> str "a4"
        | OtherPaperSize code -> obj [ req "other" (jint code) ] :> JsonNode

    let private toPrintScaling (o: JsonObject) : PrintScaling =
        match tryGet "percent" o with
        | Some n -> ScalePercent(n.GetValue<int>())
        | None ->
            let f = (childObj "fitToPage" o).Value
            FitToPage(reqInt "width" f, reqInt "height" f)

    let private ofPrintScaling (s: PrintScaling) : JsonNode =
        match s with
        | ScalePercent p -> obj [ req "percent" (jint p) ] :> JsonNode
        | FitToPage(w, h) -> obj [ req "fitToPage" (obj [ req "width" (jint w); req "height" (jint h) ]) ] :> JsonNode

    let private toPageMargins (o: JsonObject) : PageMargins =
        { Left = reqNum "left" o
          Right = reqNum "right" o
          Top = reqNum "top" o
          Bottom = reqNum "bottom" o
          Header = reqNum "header" o
          Footer = reqNum "footer" o }

    let private ofPageMargins (m: PageMargins) : JsonNode =
        obj
            [ req "left" (num m.Left)
              req "right" (num m.Right)
              req "top" (num m.Top)
              req "bottom" (num m.Bottom)
              req "header" (num m.Header)
              req "footer" (num m.Footer) ]
        :> JsonNode

    let private toPrintAreaRange (o: JsonObject) : CellRef * CellRef =
        toCellRef (reqStr "topLeft" o), toCellRef (reqStr "bottomRight" o)

    let private ofPrintAreaRange (topLeft: CellRef, bottomRight: CellRef) : JsonNode =
        obj [ req "topLeft" (str (ofCellRef topLeft)); req "bottomRight" (str (ofCellRef bottomRight)) ] :> JsonNode

    let private toPageSetup (o: JsonObject) : PageSetup =
        { Orientation = optStr "orientation" o |> Option.map toPageOrientation |> Option.defaultValue Portrait
          PaperSize = tryGet "paperSize" o |> Option.map toPaperSize
          Scaling = childObj "scaling" o |> Option.map toPrintScaling
          Margins = childObj "margins" o |> Option.map toPageMargins |> Option.defaultValue PageMargins.Default
          PrintArea =
            childArr "printArea" o
            |> Option.map (Seq.map (fun n -> toPrintAreaRange (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Header = optStr "header" o
          Footer = optStr "footer" o
          EvenHeader = optStr "evenHeader" o
          EvenFooter = optStr "evenFooter" o
          FirstHeader = optStr "firstHeader" o
          FirstFooter = optStr "firstFooter" o }

    let private ofPageSetup (p: PageSetup) : JsonNode =
        let printArea = if p.PrintArea.IsEmpty then None else Some(jarr (p.PrintArea |> List.map ofPrintAreaRange))

        obj
            [ req "orientation" (str (ofPageOrientation p.Orientation))
              optNode "paperSize" (p.PaperSize |> Option.map ofPaperSize)
              optNode "scaling" (p.Scaling |> Option.map ofPrintScaling)
              req "margins" (ofPageMargins p.Margins)
              optNode "printArea" printArea
              optNode "header" (p.Header |> Option.map str)
              optNode "footer" (p.Footer |> Option.map str)
              optNode "evenHeader" (p.EvenHeader |> Option.map str)
              optNode "evenFooter" (p.EvenFooter |> Option.map str)
              optNode "firstHeader" (p.FirstHeader |> Option.map str)
              optNode "firstFooter" (p.FirstFooter |> Option.map str) ]
        :> JsonNode

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

    let private toImageEntry (o: JsonObject) : ImageEntry =
        { Data = Convert.FromBase64String(reqStr "data" o)
          Format = toImageFormat (reqStr "format" o)
          TopLeftAnchor = toCellRef (reqStr "topLeft" o)
          BottomRightAnchor = toCellRef (reqStr "bottomRight" o) }

    let private ofImageEntry (i: ImageEntry) : JsonNode =
        obj
            [ req "format" (str (ofImageFormat i.Format))
              req "topLeft" (str (ofCellRef i.TopLeftAnchor))
              req "bottomRight" (str (ofCellRef i.BottomRightAnchor))
              req "data" (str (Convert.ToBase64String i.Data)) ]
        :> JsonNode

    // --- Tables --------------------------------------------------------------------------

    let private toTableColumn (o: JsonObject) : TableColumn =
        { Name = reqStr "name" o
          CalculatedFormula = optStr "calculatedFormula" o }

    let private ofTableColumn (c: TableColumn) : JsonNode =
        obj [ req "name" (str c.Name); optNode "calculatedFormula" (c.CalculatedFormula |> Option.map str) ] :> JsonNode

    let private toTableStyle (o: JsonObject) : TableStyle =
        { Name = optStr "name" o
          ShowFirstColumn = optBool "showFirstColumn" false o
          ShowLastColumn = optBool "showLastColumn" false o
          ShowRowStripes = optBool "showRowStripes" false o
          ShowColumnStripes = optBool "showColumnStripes" false o }

    let private ofTableStyle (s: TableStyle) : JsonNode =
        obj
            [ optNode "name" (s.Name |> Option.map str)
              (if s.ShowFirstColumn then req "showFirstColumn" (jbool true) else "showFirstColumn", None)
              (if s.ShowLastColumn then req "showLastColumn" (jbool true) else "showLastColumn", None)
              (if s.ShowRowStripes then req "showRowStripes" (jbool true) else "showRowStripes", None)
              (if s.ShowColumnStripes then req "showColumnStripes" (jbool true) else "showColumnStripes", None) ]
        :> JsonNode

    let private toTableEntry (o: JsonObject) : TableEntry =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o)
          Name = reqStr "name" o
          Columns =
            childArr "columns" o
            |> Option.map (Seq.map (fun n -> toTableColumn (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Style = childObj "style" o |> Option.map toTableStyle |> Option.defaultValue TableStyle.Default }

    let private ofTableEntry (t: TableEntry) : JsonNode =
        obj
            [ req "topLeft" (str (ofCellRef t.TopLeft))
              req "bottomRight" (str (ofCellRef t.BottomRight))
              req "name" (str t.Name)
              req "columns" (jarr (t.Columns |> List.map ofTableColumn))
              req "style" (ofTableStyle t.Style) ]
        :> JsonNode

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

    let private toSparklineStyle (o: JsonObject) : SparklineStyle =
        { Type = toSparklineType (reqStr "type" o)
          Color = tryGet "color" o |> Option.map (fun n -> toColor (n.AsObject()))
          LineWeight = optNum "lineWeight" o
          ShowMarkers = optBool "showMarkers" false o
          ShowHigh = optBool "showHigh" false o
          ShowLow = optBool "showLow" false o
          ShowFirst = optBool "showFirst" false o
          ShowLast = optBool "showLast" false o
          ShowNegative = optBool "showNegative" false o }

    let private ofSparklineStyle (s: SparklineStyle) : JsonNode =
        obj
            [ req "type" (str (ofSparklineType s.Type))
              optNode "lineWeight" (s.LineWeight |> Option.map num)
              (if s.ShowMarkers then req "showMarkers" (jbool true) else "showMarkers", None)
              (if s.ShowHigh then req "showHigh" (jbool true) else "showHigh", None)
              (if s.ShowLow then req "showLow" (jbool true) else "showLow", None)
              (if s.ShowFirst then req "showFirst" (jbool true) else "showFirst", None)
              (if s.ShowLast then req "showLast" (jbool true) else "showLast", None)
              (if s.ShowNegative then req "showNegative" (jbool true) else "showNegative", None)
              optNode "color" (s.Color |> Option.map ofColor) ]
        :> JsonNode

    let private toSparklineCell (o: JsonObject) : SparklineCell =
        { Cell = toCellRef (reqStr "cell" o)
          DataTopLeft = toCellRef (reqStr "dataTopLeft" o)
          DataBottomRight = toCellRef (reqStr "dataBottomRight" o) }

    let private ofSparklineCell (s: SparklineCell) : JsonNode =
        obj
            [ req "cell" (str (ofCellRef s.Cell))
              req "dataTopLeft" (str (ofCellRef s.DataTopLeft))
              req "dataBottomRight" (str (ofCellRef s.DataBottomRight)) ]
        :> JsonNode

    let private toSparklineGroupEntry (o: JsonObject) : SparklineGroupEntry =
        { Style = childObj "style" o |> Option.map toSparklineStyle |> Option.defaultValue SparklineStyle.Default
          Sparklines =
            childArr "sparklines" o
            |> Option.map (Seq.map (fun n -> toSparklineCell (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue [] }

    let private ofSparklineGroupEntry (g: SparklineGroupEntry) : JsonNode =
        obj
            [ req "style" (ofSparklineStyle g.Style)
              req "sparklines" (jarr (g.Sparklines |> List.map ofSparklineCell)) ]
        :> JsonNode

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

    let private toChartSeries (o: JsonObject) : ChartSeries =
        { Name = toCellRef (reqStr "name" o)
          ValuesTopLeft = toCellRef (reqStr "valuesTopLeft" o)
          ValuesBottomRight = toCellRef (reqStr "valuesBottomRight" o) }

    let private ofChartSeries (s: ChartSeries) : JsonNode =
        obj
            [ req "name" (str (ofCellRef s.Name))
              req "valuesTopLeft" (str (ofCellRef s.ValuesTopLeft))
              req "valuesBottomRight" (str (ofCellRef s.ValuesBottomRight)) ]
        :> JsonNode

    let private toChartEntry (o: JsonObject) : ChartEntry =
        let categories =
            match childObj "categories" o with
            | Some c -> c
            | None -> failwith "chart is missing required 'categories'"

        { Type = toChartType (reqStr "type" o)
          Title = optStr "title" o
          CategoriesTopLeft = toCellRef (reqStr "topLeft" categories)
          CategoriesBottomRight = toCellRef (reqStr "bottomRight" categories)
          Series =
            childArr "series" o
            |> Option.map (Seq.map (fun n -> toChartSeries (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          ShowLegend = optBool "showLegend" false o
          TopLeftAnchor = toCellRef (reqStr "anchorTopLeft" o)
          BottomRightAnchor = toCellRef (reqStr "anchorBottomRight" o) }

    let private ofChartEntry (c: ChartEntry) : JsonNode =
        obj
            [ req "type" (str (ofChartType c.Type))
              optNode "title" (c.Title |> Option.map str)
              (if c.ShowLegend then req "showLegend" (jbool true) else "showLegend", None)
              req "anchorTopLeft" (str (ofCellRef c.TopLeftAnchor))
              req "anchorBottomRight" (str (ofCellRef c.BottomRightAnchor))
              req
                  "categories"
                  (obj [ req "topLeft" (str (ofCellRef c.CategoriesTopLeft)); req "bottomRight" (str (ofCellRef c.CategoriesBottomRight)) ])
              req "series" (jarr (c.Series |> List.map ofChartSeries)) ]
        :> JsonNode

    // --- PivotTables ---------------------------------------------------------------------

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

    let private toPivotTableEntry (o: JsonObject) : PivotTableEntry =
        { SourceSheet = optStr "sourceSheet" o
          SourceTopLeft = toCellRef (reqStr "sourceTopLeft" o)
          SourceBottomRight = toCellRef (reqStr "sourceBottomRight" o)
          RowField = reqStr "rowField" o
          ColumnField = optStr "columnField" o
          ValueField = reqStr "valueField" o
          Aggregation = toPivotAggregation (reqStr "aggregation" o)
          ValueCaption = optStr "valueCaption" o
          TopLeftAnchor = toCellRef (reqStr "anchorTopLeft" o) }

    let private ofPivotTableEntry (p: PivotTableEntry) : JsonNode =
        obj
            [ optNode "sourceSheet" (p.SourceSheet |> Option.map str)
              req "sourceTopLeft" (str (ofCellRef p.SourceTopLeft))
              req "sourceBottomRight" (str (ofCellRef p.SourceBottomRight))
              req "rowField" (str p.RowField)
              optNode "columnField" (p.ColumnField |> Option.map str)
              req "valueField" (str p.ValueField)
              req "aggregation" (str (ofPivotAggregation p.Aggregation))
              optNode "valueCaption" (p.ValueCaption |> Option.map str)
              req "anchorTopLeft" (str (ofCellRef p.TopLeftAnchor)) ]
        :> JsonNode

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

    let private toConditionalFormatRule (o: JsonObject) : ConditionalFormatRule =
        match tryGet "cellValueRule" o with
        | Some n ->
            let r = n.AsObject()
            CellValueRule(toComparisonOperator (reqStr "operator" r), reqStr "formula1" r, optStr "formula2" r, toCellStyle ((childObj "style" r).Value))
        | None ->
            match tryGet "formulaRule" o with
            | Some n ->
                let r = n.AsObject()
                FormulaRule(reqStr "formula" r, toCellStyle ((childObj "style" r).Value))
            | None ->
                match tryGet "colorScale2" o with
                | Some n ->
                    let r = n.AsObject()
                    ColorScale2(toColor ((childObj "minColor" r).Value), toColor ((childObj "maxColor" r).Value))
                | None ->
                    match tryGet "colorScale3" o with
                    | Some n ->
                        let r = n.AsObject()
                        ColorScale3(toColor ((childObj "minColor" r).Value), toColor ((childObj "midColor" r).Value), toColor ((childObj "maxColor" r).Value))
                    | None ->
                        match tryGet "dataBarRule" o with
                        | Some n -> DataBarRule(toColor (n.AsObject()))
                        | None ->
                            match tryGet "duplicateValuesRule" o with
                            | Some n -> DuplicateValuesRule(toCellStyle (n.AsObject()))
                            | None ->
                                match tryGet "uniqueValuesRule" o with
                                | Some n -> UniqueValuesRule(toCellStyle (n.AsObject()))
                                | None -> failwithf "Unknown conditional format rule in %s" (o.ToJsonString())

    let private ofConditionalFormatRule (r: ConditionalFormatRule) : JsonNode =
        match r with
        | CellValueRule(op, f1, f2, style) ->
            obj
                [ req
                      "cellValueRule"
                      (obj
                          [ req "operator" (str (ofComparisonOperator op))
                            req "formula1" (str f1)
                            optNode "formula2" (f2 |> Option.map str)
                            req "style" (ofCellStyle style) ]) ]
            :> JsonNode
        | FormulaRule(formula, style) ->
            obj [ req "formulaRule" (obj [ req "formula" (str formula); req "style" (ofCellStyle style) ]) ] :> JsonNode
        | ColorScale2(minColor, maxColor) ->
            obj [ req "colorScale2" (obj [ req "minColor" (ofColor minColor); req "maxColor" (ofColor maxColor) ]) ] :> JsonNode
        | ColorScale3(minColor, midColor, maxColor) ->
            obj
                [ req
                      "colorScale3"
                      (obj [ req "minColor" (ofColor minColor); req "midColor" (ofColor midColor); req "maxColor" (ofColor maxColor) ]) ]
            :> JsonNode
        | DataBarRule color -> obj [ req "dataBarRule" (ofColor color) ] :> JsonNode
        | DuplicateValuesRule style -> obj [ req "duplicateValuesRule" (ofCellStyle style) ] :> JsonNode
        | UniqueValuesRule style -> obj [ req "uniqueValuesRule" (ofCellStyle style) ] :> JsonNode

    let private toConditionalFormatEntry (o: JsonObject) : ConditionalFormatEntry =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o)
          Rule = toConditionalFormatRule ((childObj "rule" o).Value) }

    let private ofConditionalFormatEntry (c: ConditionalFormatEntry) : JsonNode =
        obj
            [ req "topLeft" (str (ofCellRef c.TopLeft))
              req "bottomRight" (str (ofCellRef c.BottomRight))
              req "rule" (ofConditionalFormatRule c.Rule) ]
        :> JsonNode

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

    let private comparisonKindOf (o: JsonObject) : ComparisonOperator * string * string option =
        toComparisonOperator (reqStr "operator" o), reqStr "formula1" o, optStr "formula2" o

    let private ofComparisonKind (op: ComparisonOperator) (f1: string) (f2: string option) : JsonNode =
        obj [ req "operator" (str (ofComparisonOperator op)); req "formula1" (str f1); optNode "formula2" (f2 |> Option.map str) ]
        :> JsonNode

    let private toValidationKind (o: JsonObject) : ValidationKind =
        match tryGet "listValidation" o with
        | Some n -> ListValidation(n.AsArray() |> Seq.map (fun i -> i.GetValue<string>()) |> List.ofSeq)
        | None ->
            match tryGet "listFromRangeValidation" o with
            | Some n ->
                let r = n.AsObject()
                ListFromRangeValidation(toCellRef (reqStr "topLeft" r), toCellRef (reqStr "bottomRight" r))
            | None ->
                match tryGet "wholeNumberValidation" o with
                | Some n -> WholeNumberValidation(comparisonKindOf (n.AsObject()))
                | None ->
                    match tryGet "decimalValidation" o with
                    | Some n -> DecimalValidation(comparisonKindOf (n.AsObject()))
                    | None ->
                        match tryGet "textLengthValidation" o with
                        | Some n -> TextLengthValidation(comparisonKindOf (n.AsObject()))
                        | None ->
                            match tryGet "customValidation" o with
                            | Some n -> CustomValidation(n.GetValue<string>())
                            | None -> failwithf "Unknown validation kind in %s" (o.ToJsonString())

    let private ofValidationKind (k: ValidationKind) : JsonNode =
        match k with
        | ListValidation items -> obj [ req "listValidation" (jarr (items |> List.map str)) ] :> JsonNode
        | ListFromRangeValidation(topLeft, bottomRight) ->
            obj [ req "listFromRangeValidation" (obj [ req "topLeft" (str (ofCellRef topLeft)); req "bottomRight" (str (ofCellRef bottomRight)) ]) ]
            :> JsonNode
        | WholeNumberValidation(op, f1, f2) -> obj [ req "wholeNumberValidation" (ofComparisonKind op f1 f2) ] :> JsonNode
        | DecimalValidation(op, f1, f2) -> obj [ req "decimalValidation" (ofComparisonKind op f1 f2) ] :> JsonNode
        | TextLengthValidation(op, f1, f2) -> obj [ req "textLengthValidation" (ofComparisonKind op f1 f2) ] :> JsonNode
        | CustomValidation formula -> obj [ req "customValidation" (str formula) ] :> JsonNode

    let private toValidationAlert (o: JsonObject) : ValidationAlert =
        { AllowBlank = optBool "allowBlank" true o
          ErrorStyle = optStr "errorStyle" o |> Option.map toErrorAlertStyle |> Option.defaultValue Stop
          ErrorTitle = optStr "errorTitle" o
          ErrorMessage = optStr "errorMessage" o
          InputTitle = optStr "inputTitle" o
          InputMessage = optStr "inputMessage" o }

    let private ofValidationAlert (a: ValidationAlert) : JsonNode =
        obj
            [ (if not a.AllowBlank then req "allowBlank" (jbool false) else "allowBlank", None)
              (if a.ErrorStyle <> Stop then req "errorStyle" (str (ofErrorAlertStyle a.ErrorStyle)) else "errorStyle", None)
              optNode "errorTitle" (a.ErrorTitle |> Option.map str)
              optNode "errorMessage" (a.ErrorMessage |> Option.map str)
              optNode "inputTitle" (a.InputTitle |> Option.map str)
              optNode "inputMessage" (a.InputMessage |> Option.map str) ]
        :> JsonNode

    let private toDataValidationEntry (o: JsonObject) : DataValidationEntry =
        { TopLeft = toCellRef (reqStr "topLeft" o)
          BottomRight = toCellRef (reqStr "bottomRight" o)
          Kind = toValidationKind ((childObj "kind" o).Value)
          Alert = childObj "alert" o |> Option.map toValidationAlert |> Option.defaultValue ValidationAlert.Default }

    let private ofDataValidationEntry (d: DataValidationEntry) : JsonNode =
        obj
            [ req "topLeft" (str (ofCellRef d.TopLeft))
              req "bottomRight" (str (ofCellRef d.BottomRight))
              req "kind" (ofValidationKind d.Kind)
              req "alert" (ofValidationAlert d.Alert) ]
        :> JsonNode

    // --- DefinedNames --------------------------------------------------------------------

    let private toDefinedNameScope (node: JsonNode) : DefinedNameScope =
        match node with
        | :? JsonValue as v when (v.TryGetValue<string>() |> fst) && v.GetValue<string>() = "workbookScope" -> WorkbookScope
        | _ -> SheetScope(reqStr "sheetScope" (node.AsObject()))

    let private ofDefinedNameScope (s: DefinedNameScope) : JsonNode =
        match s with
        | WorkbookScope -> str "workbookScope"
        | SheetScope sheetName -> obj [ req "sheetScope" (str sheetName) ] :> JsonNode

    let private toDefinedNameEntry (o: JsonObject) : DefinedNameEntry =
        { Name = reqStr "name" o
          Formula = reqStr "formula" o
          Scope = toDefinedNameScope ((tryGet "scope" o).Value)
          Hidden = optBool "hidden" false o }

    let private ofDefinedNameEntry (d: DefinedNameEntry) : JsonNode =
        obj
            [ req "name" (str d.Name)
              req "formula" (str d.Formula)
              req "scope" (ofDefinedNameScope d.Scope)
              (if d.Hidden then req "hidden" (jbool true) else "hidden", None) ]
        :> JsonNode

    // --- Worksheet / Workbook ----------------------------------------------------------

    let private toWorksheet (o: JsonObject) : Worksheet =
        { Name = reqStr "name" o
          Cells = childArr "cells" o |> Option.map (Seq.map (fun n -> toCell (n.AsObject())) >> List.ofSeq) |> Option.defaultValue []
          ColumnProps =
            childArr "columnProps" o
            |> Option.map (Seq.map (fun n -> toColumnProps (n.AsObject())) >> Map.ofSeq)
            |> Option.defaultValue Map.empty
          RowProps =
            childArr "rowProps" o
            |> Option.map (Seq.map (fun n -> toRowProps (n.AsObject())) >> Map.ofSeq)
            |> Option.defaultValue Map.empty
          MergedRanges =
            childArr "mergedRanges" o
            |> Option.map (Seq.map (fun n -> toMergedRange (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          FreezePane = childObj "freezePane" o |> Option.map toFreezePane
          AutoFilter = childObj "autoFilter" o |> Option.map toAutoFilterRange
          Protection = childObj "protection" o |> Option.map toSheetProtection
          ConditionalFormats =
            childArr "conditionalFormats" o
            |> Option.map (Seq.map (fun n -> toConditionalFormatEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          DataValidations =
            childArr "dataValidations" o
            |> Option.map (Seq.map (fun n -> toDataValidationEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Hyperlinks =
            childArr "hyperlinks" o
            |> Option.map (Seq.map (fun n -> toHyperlinkEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Comments =
            childArr "comments" o
            |> Option.map (Seq.map (fun n -> toCommentEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          PageSetup = childObj "pageSetup" o |> Option.map toPageSetup
          Tables =
            childArr "tables" o
            |> Option.map (Seq.map (fun n -> toTableEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          SparklineGroups =
            childArr "sparklineGroups" o
            |> Option.map (Seq.map (fun n -> toSparklineGroupEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Charts =
            childArr "charts" o
            |> Option.map (Seq.map (fun n -> toChartEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Images =
            childArr "images" o
            |> Option.map (Seq.map (fun n -> toImageEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          PivotTables =
            childArr "pivotTables" o
            |> Option.map (Seq.map (fun n -> toPivotTableEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue [] }

    let private sortedByCell (refOf: 'a -> CellRef) (items: 'a list) : 'a list = items |> List.sortBy refOf

    let private ofWorksheet (s: Worksheet) : JsonNode =
        let cells = s.Cells |> sortedByCell (fun c -> c.Ref) |> List.map ofCell
        let mergedRanges = s.MergedRanges |> sortedByCell (fun m -> m.TopLeft) |> List.map ofMergedRange
        let columnProps = s.ColumnProps |> Map.toList |> List.map (fun (i, c) -> ofColumnProps i c)
        let rowProps = s.RowProps |> Map.toList |> List.map (fun (i, r) -> ofRowProps i r)
        let comments = s.Comments |> sortedByCell (fun c -> c.Cell) |> List.map ofCommentEntry
        let hyperlinks = s.Hyperlinks |> sortedByCell (fun h -> h.TopLeft) |> List.map ofHyperlinkEntry
        let images = s.Images |> sortedByCell (fun i -> i.TopLeftAnchor) |> List.map ofImageEntry
        let tables = s.Tables |> sortedByCell (fun t -> t.TopLeft) |> List.map ofTableEntry

        let sparklineGroupAnchor (g: SparklineGroupEntry) =
            g.Sparklines |> List.map (fun sl -> sl.Cell) |> List.sortBy id |> List.tryHead |> Option.defaultValue (CellRef.create 0 0)

        let sparklineGroups = s.SparklineGroups |> List.sortBy sparklineGroupAnchor |> List.map ofSparklineGroupEntry
        let charts = s.Charts |> sortedByCell (fun c -> c.TopLeftAnchor) |> List.map ofChartEntry
        let pivotTables = s.PivotTables |> sortedByCell (fun p -> p.TopLeftAnchor) |> List.map ofPivotTableEntry
        let conditionalFormats = s.ConditionalFormats |> sortedByCell (fun c -> c.TopLeft) |> List.map ofConditionalFormatEntry
        let dataValidations = s.DataValidations |> sortedByCell (fun d -> d.TopLeft) |> List.map ofDataValidationEntry

        obj
            [ req "name" (str s.Name)
              (if cells.IsEmpty then "cells", None else req "cells" (jarr cells))
              (if mergedRanges.IsEmpty then "mergedRanges", None else req "mergedRanges" (jarr mergedRanges))
              optNode "freezePane" (s.FreezePane |> Option.map ofFreezePane)
              optNode "autoFilter" (s.AutoFilter |> Option.map ofAutoFilterRange)
              (if columnProps.IsEmpty then "columnProps", None else req "columnProps" (jarr columnProps))
              (if rowProps.IsEmpty then "rowProps", None else req "rowProps" (jarr rowProps))
              (if comments.IsEmpty then "comments", None else req "comments" (jarr comments))
              (if hyperlinks.IsEmpty then "hyperlinks", None else req "hyperlinks" (jarr hyperlinks))
              optNode "protection" (s.Protection |> Option.map ofSheetProtection)
              optNode "pageSetup" (s.PageSetup |> Option.map ofPageSetup)
              (if images.IsEmpty then "images", None else req "images" (jarr images))
              (if tables.IsEmpty then "tables", None else req "tables" (jarr tables))
              (if sparklineGroups.IsEmpty then "sparklineGroups", None else req "sparklineGroups" (jarr sparklineGroups))
              (if charts.IsEmpty then "charts", None else req "charts" (jarr charts))
              (if pivotTables.IsEmpty then "pivotTables", None else req "pivotTables" (jarr pivotTables))
              (if conditionalFormats.IsEmpty then
                   "conditionalFormats", None
               else
                   req "conditionalFormats" (jarr conditionalFormats))
              (if dataValidations.IsEmpty then
                   "dataValidations", None
               else
                   req "dataValidations" (jarr dataValidations)) ]
        :> JsonNode

    /// Reads a `Workbook` from a JSON tree. `root` should be the top-level `{"sheets": [...]
    /// ...}` object - see this module's own doc comment for the schema `ofWorkbook`
    /// produces (the two are each other's exact inverse for anything this pass covers).
    let toWorkbook (root: JsonObject) : Workbook =
        { Sheets = childArr "sheets" root |> Option.map (Seq.map (fun n -> toWorksheet (n.AsObject())) >> List.ofSeq) |> Option.defaultValue []
          DefinedNames =
            childArr "definedNames" root
            |> Option.map (Seq.map (fun n -> toDefinedNameEntry (n.AsObject())) >> List.ofSeq)
            |> Option.defaultValue []
          Protection = childObj "protection" root |> Option.map toWorkbookProtection
          VbaProject = optStr "vbaProject" root |> Option.map Convert.FromBase64String }

    /// Renders a `Workbook` as a JSON tree rooted at a plain `{"sheets": [...], ...}` object.
    let ofWorkbook (wb: Workbook) : JsonObject =
        let definedNames = wb.DefinedNames |> List.sortBy (fun d -> d.Name) |> List.map ofDefinedNameEntry

        obj
            [ req "sheets" (jarr (wb.Sheets |> List.map ofWorksheet))
              (if definedNames.IsEmpty then "definedNames", None else req "definedNames" (jarr definedNames))
              optNode "protection" (wb.Protection |> Option.map ofWorkbookProtection)
              optNode "vbaProject" (wb.VbaProject |> Option.map (Convert.ToBase64String >> str)) ]
