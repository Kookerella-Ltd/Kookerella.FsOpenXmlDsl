namespace SafeOpenXml.Interpreter

open System.Collections.Generic
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

module internal ColorMapping =

    let private hex (r: byte) (g: byte) (b: byte) =
        HexBinaryValue(sprintf "FF%02X%02X%02X" r g b)

    let colorElement (color: Color) : Spreadsheet.Color =
        match color with
        | Rgb(r, g, b) -> Spreadsheet.Color(Rgb = hex r g b)
        | Indexed i -> Spreadsheet.Color(Indexed = UInt32Value(uint32 i))
        | Theme(idx, tint) ->
            let c = Spreadsheet.Color(Theme = UInt32Value(uint32 idx))
            tint |> Option.iter (fun t -> c.Tint <- DoubleValue(t))
            c

    let foregroundColor (color: Color) : ForegroundColor =
        match color with
        | Rgb(r, g, b) -> ForegroundColor(Rgb = hex r g b)
        | Indexed i -> ForegroundColor(Indexed = UInt32Value(uint32 i))
        | Theme(idx, tint) ->
            let c = ForegroundColor(Theme = UInt32Value(uint32 idx))
            tint |> Option.iter (fun t -> c.Tint <- DoubleValue(t))
            c

module internal BorderMapping =

    let toOpenXml (style: BorderLineStyle) : BorderStyleValues =
        match style with
        | Thin -> BorderStyleValues.Thin
        | Medium -> BorderStyleValues.Medium
        | Thick -> BorderStyleValues.Thick
        | Dashed -> BorderStyleValues.Dashed
        | Dotted -> BorderStyleValues.Dotted
        | Double -> BorderStyleValues.Double
        | Hair -> BorderStyleValues.Hair
        | Other raw -> BorderStyleValues(raw)

    let ofOpenXml (value: BorderStyleValues) : BorderLineStyle =
        if value = BorderStyleValues.Thin then Thin
        elif value = BorderStyleValues.Medium then Medium
        elif value = BorderStyleValues.Thick then Thick
        elif value = BorderStyleValues.Dashed then Dashed
        elif value = BorderStyleValues.Dotted then Dotted
        elif value = BorderStyleValues.Double then Double
        elif value = BorderStyleValues.Hair then Hair
        else Other(value.ToString())

module internal AlignmentMapping =

    let horizontalToOpenXml (h: HorizontalAlignment) : HorizontalAlignmentValues =
        match h with
        | GeneralAlign -> HorizontalAlignmentValues.General
        | AlignLeft -> HorizontalAlignmentValues.Left
        | AlignCenter -> HorizontalAlignmentValues.Center
        | AlignRight -> HorizontalAlignmentValues.Right
        | AlignFill -> HorizontalAlignmentValues.Fill
        | AlignJustify -> HorizontalAlignmentValues.Justify

    let horizontalOfOpenXml (v: HorizontalAlignmentValues) : HorizontalAlignment =
        if v = HorizontalAlignmentValues.Left then AlignLeft
        elif v = HorizontalAlignmentValues.Center then AlignCenter
        elif v = HorizontalAlignmentValues.Right then AlignRight
        elif v = HorizontalAlignmentValues.Fill then AlignFill
        elif v = HorizontalAlignmentValues.Justify then AlignJustify
        else GeneralAlign

    let verticalToOpenXml (v: VerticalAlignment) : VerticalAlignmentValues =
        match v with
        | AlignTop -> VerticalAlignmentValues.Top
        | AlignMiddle -> VerticalAlignmentValues.Center
        | AlignBottom -> VerticalAlignmentValues.Bottom

    let verticalOfOpenXml (v: VerticalAlignmentValues) : VerticalAlignment =
        if v = VerticalAlignmentValues.Top then AlignTop
        elif v = VerticalAlignmentValues.Bottom then AlignBottom
        else AlignMiddle

module internal NumberFormatMapping =

    // Built-in OOXML numFmtIds. See ECMA-376 Part 1, §18.8.30.
    let [<Literal>] GeneralId = 0u
    let [<Literal>] IntegerId = 1u
    let [<Literal>] TwoDecimalId = 2u
    let [<Literal>] PercentageId = 10u
    let [<Literal>] ShortDateId = 14u
    let [<Literal>] DateTimeId = 22u
    let [<Literal>] CurrencyFormatCode = "\"$\"#,##0.00"

/// Interns fonts/fills/borders/number formats into shared stylesheet entries, mirroring
/// how Excel itself deduplicates styles - identical `CellStyle` values (structural
/// equality, for free on an F# record) always resolve to the same cell format index.
type internal StyleRegistry() =
    let fontList = ResizeArray<FontStyle>()
    let fontIndex = Dictionary<FontStyle, uint32>()
    let fillList = ResizeArray<FillStyle>()
    let fillIndex = Dictionary<FillStyle, uint32>()
    let borderList = ResizeArray<BorderStyle>()
    let borderIndex = Dictionary<BorderStyle, uint32>()
    let customNumFmts = Dictionary<string, uint32>()
    let mutable nextCustomNumFmtId = 164u
    // (fontId, fillId, borderId, numFmtId, alignment)
    let cellFormats = ResizeArray<uint32 * uint32 * uint32 * uint32 * AlignmentStyle option>()
    let cellFormatIndex = Dictionary<CellStyle, uint32>()

    do
        fontList.Add(FontStyle.Default)
        fontIndex.[FontStyle.Default] <- 0u
        borderList.Add(BorderStyle.None)
        borderIndex.[BorderStyle.None] <- 0u
        cellFormats.Add(0u, 0u, 0u, 0u, None)
        cellFormatIndex.[CellStyle.Default] <- 0u

    member private _.InternFont(fontOpt: FontStyle option) : uint32 =
        match fontOpt with
        | None -> 0u
        | Some f ->
            match fontIndex.TryGetValue f with
            | true, idx -> idx
            | false, _ ->
                let idx = uint32 fontList.Count
                fontList.Add f
                fontIndex.[f] <- idx
                idx

    member private _.InternFill(fillOpt: FillStyle option) : uint32 =
        match fillOpt with
        | None -> 0u
        | Some fill ->
            match fillIndex.TryGetValue fill with
            | true, idx -> idx
            | false, _ ->
                // Offset by 2: slot 0 is the mandatory "none" fill, slot 1 is "gray125".
                let idx = uint32 (fillList.Count + 2)
                fillList.Add fill
                fillIndex.[fill] <- idx
                idx

    member private _.InternBorder(borderOpt: BorderStyle option) : uint32 =
        let border = defaultArg borderOpt BorderStyle.None
        match borderIndex.TryGetValue border with
        | true, idx -> idx
        | false, _ ->
            let idx = uint32 borderList.Count
            borderList.Add border
            borderIndex.[border] <- idx
            idx

    member private _.InternCustomFormat(code: string) : uint32 =
        match customNumFmts.TryGetValue code with
        | true, id -> id
        | false, _ ->
            let id = nextCustomNumFmtId
            nextCustomNumFmtId <- nextCustomNumFmtId + 1u
            customNumFmts.[code] <- id
            id

    member this.InternNumberFormat(nfOpt: NumberFormat option) : uint32 =
        match nfOpt with
        | None -> NumberFormatMapping.GeneralId
        | Some General -> NumberFormatMapping.GeneralId
        | Some Integer -> NumberFormatMapping.IntegerId
        | Some TwoDecimal -> NumberFormatMapping.TwoDecimalId
        | Some Percentage -> NumberFormatMapping.PercentageId
        | Some Currency -> this.InternCustomFormat NumberFormatMapping.CurrencyFormatCode
        | Some ShortDate -> NumberFormatMapping.ShortDateId
        | Some DateAndTime -> NumberFormatMapping.DateTimeId
        | Some(Custom code) -> this.InternCustomFormat code

    /// Returns the CellFormat ("style") index to put in a cell's `s` attribute.
    member this.GetCellFormatIndex(styleOpt: CellStyle option) : uint32 =
        let style = defaultArg styleOpt CellStyle.Default
        match cellFormatIndex.TryGetValue style with
        | true, idx -> idx
        | false, _ ->
            let fontId = this.InternFont style.Font
            let fillId = this.InternFill style.Fill
            let borderId = this.InternBorder style.Border
            let numFmtId = this.InternNumberFormat style.NumberFormat
            let idx = uint32 cellFormats.Count
            cellFormats.Add(fontId, fillId, borderId, numFmtId, style.Alignment)
            cellFormatIndex.[style] <- idx
            idx

    member private _.FontToOpenXml(f: FontStyle) : Font =
        let font = Font()
        f.Name |> Option.iter (fun n -> font.AppendChild(FontName(Val = StringValue(n))) |> ignore)
        f.Size |> Option.iter (fun s -> font.AppendChild(FontSize(Val = DoubleValue(s))) |> ignore)
        if f.Bold then font.AppendChild(Bold()) |> ignore
        if f.Italic then font.AppendChild(Italic()) |> ignore
        if f.Underline then font.AppendChild(Underline()) |> ignore
        if f.Strikethrough then font.AppendChild(Strike()) |> ignore
        f.Color |> Option.iter (fun c -> font.AppendChild(ColorMapping.colorElement c) |> ignore)
        font

    member private _.FillToOpenXml(fill: FillStyle) : Fill =
        let patternFill = PatternFill(PatternType = EnumValue<PatternValues>(PatternValues.Solid))
        // OOXML quirk: for a solid pattern fill, the color Excel shows as the cell's fill
        // color is the *foreground* color, not the background color.
        patternFill.AppendChild(ColorMapping.foregroundColor fill.Color) |> ignore
        let fillEl = Fill()
        fillEl.AppendChild(patternFill) |> ignore
        fillEl

    member private _.BorderToOpenXml(b: BorderStyle) : Border =
        let border = Border()

        let appendSide
            (sideOpt: BorderSide option)
            (make: unit -> 'T)
            (setStyle: 'T -> EnumValue<BorderStyleValues> -> unit)
            (appendColor: 'T -> Spreadsheet.Color -> unit)
            (appendChild: 'T -> unit)
            =
            match sideOpt with
            | Some side ->
                let el = make ()
                setStyle el (EnumValue<BorderStyleValues>(BorderMapping.toOpenXml side.Style))
                side.Color |> Option.iter (fun c -> appendColor el (ColorMapping.colorElement c))
                appendChild el
            | None -> ()

        appendSide
            b.Left
            (fun () -> LeftBorder())
            (fun el v -> el.Style <- v)
            (fun el c -> el.AppendChild(c) |> ignore)
            (fun el -> border.AppendChild(el) |> ignore)

        appendSide
            b.Right
            (fun () -> RightBorder())
            (fun el v -> el.Style <- v)
            (fun el c -> el.AppendChild(c) |> ignore)
            (fun el -> border.AppendChild(el) |> ignore)

        appendSide
            b.Top
            (fun () -> TopBorder())
            (fun el v -> el.Style <- v)
            (fun el c -> el.AppendChild(c) |> ignore)
            (fun el -> border.AppendChild(el) |> ignore)

        appendSide
            b.Bottom
            (fun () -> BottomBorder())
            (fun el v -> el.Style <- v)
            (fun el c -> el.AppendChild(c) |> ignore)
            (fun el -> border.AppendChild(el) |> ignore)

        border

    member private _.CellFormatToOpenXml
        (fontId: uint32, fillId: uint32, borderId: uint32, numFmtId: uint32, alignment: AlignmentStyle option)
        : CellFormat =
        let cf =
            CellFormat(
                FontId = UInt32Value(fontId),
                FillId = UInt32Value(fillId),
                BorderId = UInt32Value(borderId),
                NumberFormatId = UInt32Value(numFmtId)
            )

        if fontId <> 0u then cf.ApplyFont <- BooleanValue(true)
        if fillId <> 0u then cf.ApplyFill <- BooleanValue(true)
        if borderId <> 0u then cf.ApplyBorder <- BooleanValue(true)
        if numFmtId <> 0u then cf.ApplyNumberFormat <- BooleanValue(true)

        match alignment with
        | Some a ->
            cf.ApplyAlignment <- BooleanValue(true)
            let al = Alignment()

            a.Horizontal
            |> Option.iter (fun h ->
                al.Horizontal <- EnumValue<HorizontalAlignmentValues>(AlignmentMapping.horizontalToOpenXml h))

            a.Vertical
            |> Option.iter (fun v ->
                al.Vertical <- EnumValue<VerticalAlignmentValues>(AlignmentMapping.verticalToOpenXml v))

            if a.WrapText then
                al.WrapText <- BooleanValue(true)

            cf.AppendChild(al) |> ignore
        | None -> ()

        cf

    /// Assembles everything interned so far into a single OOXML `Stylesheet`.
    member this.BuildStylesheet() : Stylesheet =
        let fonts = Fonts(Count = UInt32Value(uint32 fontList.Count))
        fontList |> Seq.iter (fun f -> fonts.AppendChild(this.FontToOpenXml f) |> ignore)

        let noFill = Fill()
        noFill.AppendChild(PatternFill(PatternType = EnumValue<PatternValues>(PatternValues.None))) |> ignore
        let gray125Fill = Fill()
        gray125Fill.AppendChild(PatternFill(PatternType = EnumValue<PatternValues>(PatternValues.Gray125))) |> ignore
        let fills = Fills(Count = UInt32Value(uint32 (fillList.Count + 2)))
        fills.AppendChild(noFill) |> ignore
        fills.AppendChild(gray125Fill) |> ignore
        fillList |> Seq.iter (fun f -> fills.AppendChild(this.FillToOpenXml f) |> ignore)

        let borders = Borders(Count = UInt32Value(uint32 borderList.Count))
        borderList |> Seq.iter (fun b -> borders.AppendChild(this.BorderToOpenXml b) |> ignore)

        let stylesheet = Stylesheet()

        if customNumFmts.Count > 0 then
            let numFmts = NumberingFormats(Count = UInt32Value(uint32 customNumFmts.Count))

            customNumFmts
            |> Seq.iter (fun kv ->
                numFmts.AppendChild(NumberingFormat(NumberFormatId = UInt32Value(kv.Value), FormatCode = StringValue(kv.Key)))
                |> ignore)

            stylesheet.AppendChild(numFmts) |> ignore

        stylesheet.AppendChild(fonts) |> ignore
        stylesheet.AppendChild(fills) |> ignore
        stylesheet.AppendChild(borders) |> ignore

        let cellStyleFormats = CellStyleFormats(Count = UInt32Value(1u))

        cellStyleFormats.AppendChild(
            CellFormat(FontId = UInt32Value(0u), FillId = UInt32Value(0u), BorderId = UInt32Value(0u), NumberFormatId = UInt32Value(0u))
        )
        |> ignore

        stylesheet.AppendChild(cellStyleFormats) |> ignore

        let cellFormatsEl = CellFormats(Count = UInt32Value(uint32 cellFormats.Count))
        cellFormats |> Seq.iter (fun cf -> cellFormatsEl.AppendChild(this.CellFormatToOpenXml cf) |> ignore)
        stylesheet.AppendChild(cellFormatsEl) |> ignore

        let cellStyles = CellStyles(Count = UInt32Value(1u))

        cellStyles.AppendChild(Spreadsheet.CellStyle(Name = StringValue("Normal"), FormatId = UInt32Value(0u), BuiltinId = UInt32Value(0u)))
        |> ignore

        stylesheet.AppendChild(cellStyles) |> ignore

        stylesheet
