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
/// sizing) plus VBA and defined names - see this module's own progress in `CLAUDE.md`/the
/// project README for what's covered so far.
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
          Protection = None
          ConditionalFormats = []
          DataValidations = []
          Hyperlinks = []
          Comments = []
          PageSetup = None
          Tables = []
          SparklineGroups = []
          Charts = []
          Images = []
          PivotTables = [] }

    let private sortedByCell (refOf: 'a -> CellRef) (items: 'a list) : 'a list = items |> List.sortBy refOf

    let private ofWorksheet (s: Worksheet) : JsonNode =
        let cells = s.Cells |> sortedByCell (fun c -> c.Ref) |> List.map ofCell
        let mergedRanges = s.MergedRanges |> sortedByCell (fun m -> m.TopLeft) |> List.map ofMergedRange
        let columnProps = s.ColumnProps |> Map.toList |> List.map (fun (i, c) -> ofColumnProps i c)
        let rowProps = s.RowProps |> Map.toList |> List.map (fun (i, r) -> ofRowProps i r)

        obj
            [ req "name" (str s.Name)
              (if cells.IsEmpty then "cells", None else req "cells" (jarr cells))
              (if mergedRanges.IsEmpty then "mergedRanges", None else req "mergedRanges" (jarr mergedRanges))
              optNode "freezePane" (s.FreezePane |> Option.map ofFreezePane)
              optNode "autoFilter" (s.AutoFilter |> Option.map ofAutoFilterRange)
              (if columnProps.IsEmpty then "columnProps", None else req "columnProps" (jarr columnProps))
              (if rowProps.IsEmpty then "rowProps", None else req "rowProps" (jarr rowProps)) ]
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
          Protection = None
          VbaProject = optStr "vbaProject" root |> Option.map Convert.FromBase64String }

    /// Renders a `Workbook` as a JSON tree rooted at a plain `{"sheets": [...], ...}` object.
    let ofWorkbook (wb: Workbook) : JsonObject =
        let definedNames = wb.DefinedNames |> List.sortBy (fun d -> d.Name) |> List.map ofDefinedNameEntry

        obj
            [ req "sheets" (jarr (wb.Sheets |> List.map ofWorksheet))
              (if definedNames.IsEmpty then "definedNames", None else req "definedNames" (jarr definedNames))
              optNode "vbaProject" (wb.VbaProject |> Option.map (Convert.ToBase64String >> str)) ]
