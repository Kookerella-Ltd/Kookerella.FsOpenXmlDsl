namespace SafeOpenXml.Interpreter

open System
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// Parses the DrawingML/ChartML parts anchored on one worksheet back into `ChartEntry`
/// values - the read side of `ChartWriter`. See that module's own doc comment for why
/// this is a separate file, and for why none of `DocumentFormat.OpenXml.Drawing`/
/// `.Drawing.Charts`/`.Drawing.Spreadsheet` are `open`ed.
module internal ChartReader =

    type DChart = DocumentFormat.OpenXml.Drawing.Charts.Chart
    type DChartSpace = DocumentFormat.OpenXml.Drawing.Charts.ChartSpace
    type DPlotArea = DocumentFormat.OpenXml.Drawing.Charts.PlotArea
    type DBarChart = DocumentFormat.OpenXml.Drawing.Charts.BarChart
    type DBarChartSeries = DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries
    type DLineChart = DocumentFormat.OpenXml.Drawing.Charts.LineChart
    type DLineChartSeries = DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries
    type DPieChart = DocumentFormat.OpenXml.Drawing.Charts.PieChart
    type DPieChartSeries = DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries
    type DCategoryAxisData = DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData
    type DValues = DocumentFormat.OpenXml.Drawing.Charts.Values
    type DSeriesText = DocumentFormat.OpenXml.Drawing.Charts.SeriesText
    type DBarDirectionValues = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues

    type XTwoCellAnchor = DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor
    type XGraphicFrame = DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame

    type XChartReference = DocumentFormat.OpenXml.Drawing.Charts.ChartReference

    let private stripSheetQualifier (s: string) =
        match s.LastIndexOf('!') with
        | -1 -> s
        | i -> s.Substring(i + 1)

    let private parseSingleCellFormula (formula: string) : CellRef =
        CellRef.ofA1 ((stripSheetQualifier formula).Replace("$", ""))

    let private parseRangeFormula (formula: string) : CellRef * CellRef =
        let refPart = (stripSheetQualifier formula).Replace("$", "")
        let parts = refPart.Split(':')
        let topLeft = CellRef.ofA1 parts.[0]
        let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])
        (topLeft, bottomRight)

    let private nameCellOf (seriesText: DSeriesText) : CellRef option =
        Option.ofObj seriesText
        |> Option.bind (fun st -> Option.ofObj st.StringReference)
        |> Option.bind (fun sr -> Option.ofObj sr.Formula)
        |> Option.map (fun f -> parseSingleCellFormula f.Text)

    let private categoriesOf (cat: DCategoryAxisData) : (CellRef * CellRef) option =
        Option.ofObj cat
        |> Option.bind (fun c -> Option.ofObj c.StringReference)
        |> Option.bind (fun sr -> Option.ofObj sr.Formula)
        |> Option.map (fun f -> parseRangeFormula f.Text)

    let private valuesRangeOf (vals: DValues) : (CellRef * CellRef) option =
        Option.ofObj vals
        |> Option.bind (fun v -> Option.ofObj v.NumberReference)
        |> Option.bind (fun nr -> Option.ofObj nr.Formula)
        |> Option.map (fun f -> parseRangeFormula f.Text)

    /// `Name`/`Categories`/`Values` are each independently optional depending on what's
    /// present in the file - `chartEntryFromXxx` below is what decides a series without a
    /// usable name/values, or a chart with no usable categories at all, gets dropped
    /// rather than producing a `ChartSeries` with made-up data.
    let private barSeriesOf (ser: DBarChartSeries) : CellRef option * (CellRef * CellRef) option * (CellRef * CellRef) option =
        let name = nameCellOf ser.SeriesText
        let cat = ser.Elements<DCategoryAxisData>() |> Seq.tryHead |> Option.bind categoriesOf
        let vals = ser.Elements<DValues>() |> Seq.tryHead |> Option.bind valuesRangeOf
        (name, cat, vals)

    let private lineSeriesOf (ser: DLineChartSeries) : CellRef option * (CellRef * CellRef) option * (CellRef * CellRef) option =
        let name = nameCellOf ser.SeriesText
        let cat = ser.Elements<DCategoryAxisData>() |> Seq.tryHead |> Option.bind categoriesOf
        let vals = ser.Elements<DValues>() |> Seq.tryHead |> Option.bind valuesRangeOf
        (name, cat, vals)

    let private pieSeriesOf (ser: DPieChartSeries) : CellRef option * (CellRef * CellRef) option * (CellRef * CellRef) option =
        let name = nameCellOf ser.SeriesText
        let cat = ser.Elements<DCategoryAxisData>() |> Seq.tryHead |> Option.bind categoriesOf
        let vals = ser.Elements<DValues>() |> Seq.tryHead |> Option.bind valuesRangeOf
        (name, cat, vals)

    let private toSeriesList (results: (CellRef option * (CellRef * CellRef) option * (CellRef * CellRef) option) list) : (CellRef * CellRef) option * ChartSeries list =
        let categories = results |> List.tryPick (fun (_, cat, _) -> cat)

        let series =
            results
            |> List.choose (fun (name, _, vals) ->
                match name, vals with
                | Some n, Some(vTL, vBR) -> Some { Name = n; ValuesTopLeft = vTL; ValuesBottomRight = vBR }
                | _ -> None)

        (categories, series)

    let private chartEntryFromBar (bar: DBarChart) : ChartType * (CellRef * CellRef) option * ChartSeries list =
        let chartType =
            match Option.ofObj bar.BarDirection |> Option.bind (fun d -> Option.ofObj d.Val) with
            | Some v when v.Value = DBarDirectionValues.Bar -> ChartBar
            | _ -> ChartColumn

        let categories, series = bar.Elements<DBarChartSeries>() |> Seq.map barSeriesOf |> List.ofSeq |> toSeriesList
        (chartType, categories, series)

    let private chartEntryFromLine (line: DLineChart) : ChartType * (CellRef * CellRef) option * ChartSeries list =
        let categories, series = line.Elements<DLineChartSeries>() |> Seq.map lineSeriesOf |> List.ofSeq |> toSeriesList
        (ChartLine, categories, series)

    let private chartEntryFromPie (pie: DPieChart) : ChartType * (CellRef * CellRef) option * ChartSeries list =
        let categories, series = pie.Elements<DPieChartSeries>() |> Seq.map pieSeriesOf |> List.ofSeq |> toSeriesList
        (ChartPie, categories, series)

    let private chartEntryOf (chartSpace: DChartSpace) (topLeftAnchor: CellRef) (bottomRightAnchor: CellRef) : ChartEntry option =
        chartSpace.Elements<DChart>()
        |> Seq.tryHead
        |> Option.bind (fun chart -> Option.ofObj chart.PlotArea |> Option.map (fun plotArea -> chart, plotArea))
        |> Option.bind (fun (chart, plotArea) ->
            let typeCatSeries =
                match plotArea.Elements<DBarChart>() |> Seq.tryHead with
                | Some bar -> Some(chartEntryFromBar bar)
                | None ->
                    match plotArea.Elements<DLineChart>() |> Seq.tryHead with
                    | Some line -> Some(chartEntryFromLine line)
                    | None -> plotArea.Elements<DPieChart>() |> Seq.tryHead |> Option.map chartEntryFromPie

            typeCatSeries |> Option.map (fun tcs -> chart, tcs))
        |> Option.bind (fun (chart, (chartType, categories, series)) ->
            categories
            |> Option.map (fun (catTL, catBR) ->
                let title =
                    if isNull chart.Title || isNull chart.Title.ChartText then
                        None
                    else
                        let text = chart.Title.ChartText.InnerText
                        if String.IsNullOrEmpty text then None else Some text

                { Type = chartType
                  Title = title
                  CategoriesTopLeft = catTL
                  CategoriesBottomRight = catBR
                  Series = series
                  ShowLegend = not (isNull chart.Legend)
                  TopLeftAnchor = topLeftAnchor
                  BottomRightAnchor = bottomRightAnchor }))

    /// Tries to interpret one anchor as a chart - `None` if it doesn't contain a
    /// `graphicFrame`/chart relationship at all (e.g. it's a `Picture` instead - see
    /// `ImageReader.tryReadImage` - or a chart this can't make sense of: no categories, no
    /// chart-type element Core recognizes). `DrawingReader` is what actually walks the
    /// worksheet's `<drawing>` relationship to find `drawingsPart`/`anchor` in the first
    /// place, since the same walk feeds both this and `ImageReader.tryReadImage`.
    let tryReadChart (drawingsPart: DrawingsPart) (anchor: XTwoCellAnchor) (topLeftAnchor: CellRef) (bottomRightAnchor: CellRef) : ChartEntry option =
        anchor.Elements<XGraphicFrame>()
        |> Seq.tryHead
        |> Option.bind (fun frame -> Option.ofObj frame.Graphic)
        |> Option.bind (fun g -> Option.ofObj g.GraphicData)
        |> Option.bind (fun gd -> gd.Elements<XChartReference>() |> Seq.tryHead)
        |> Option.bind (fun cr -> Option.ofObj cr.Id)
        |> Option.bind (fun relIdVal ->
            match drawingsPart.GetPartById(relIdVal.Value) with
            | :? ChartPart as chartPart -> Option.ofObj chartPart.ChartSpace
            | _ -> None)
        |> Option.bind (fun chartSpace -> chartEntryOf chartSpace topLeftAnchor bottomRightAnchor)
