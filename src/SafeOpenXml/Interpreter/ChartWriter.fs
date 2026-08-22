namespace SafeOpenXml.Interpreter

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// Builds the DrawingML/ChartML parts and relationships for the charts anchored on one
/// worksheet - the write side of `ChartReader`. Charts are the one feature in this
/// library spanning several OOXML parts at once (a `DrawingsPart` per worksheet, a
/// `ChartPart` per chart, referenced from the worksheet via a `drawing` element), which is
/// why this lives in its own file rather than inline in `Writer.fs` the way smaller
/// features do.
///
/// None of `DocumentFormat.OpenXml.Drawing`/`.Drawing.Charts`/`.Drawing.Spreadsheet` are
/// `open`ed - see `EmbeddedChart`'s doc comment in `Builders.fs` for why (that surface
/// collides with several already-qualified names: `Spreadsheet.PageSetup`/`.PageMargins`/
/// `.Protection`, `CellValue.Formula`, and this DSL's own `CellValue.Text` case). Every
/// type needed from them is instead an explicit abbreviation below.
module internal ChartWriter =

    type DChart = DocumentFormat.OpenXml.Drawing.Charts.Chart
    type DChartSpace = DocumentFormat.OpenXml.Drawing.Charts.ChartSpace
    type DPlotArea = DocumentFormat.OpenXml.Drawing.Charts.PlotArea
    type DBarChart = DocumentFormat.OpenXml.Drawing.Charts.BarChart
    type DBarChartSeries = DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries
    type DLineChart = DocumentFormat.OpenXml.Drawing.Charts.LineChart
    type DLineChartSeries = DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries
    type DPieChart = DocumentFormat.OpenXml.Drawing.Charts.PieChart
    type DPieChartSeries = DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries
    type DCategoryAxis = DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis
    type DValueAxis = DocumentFormat.OpenXml.Drawing.Charts.ValueAxis
    type DScaling = DocumentFormat.OpenXml.Drawing.Charts.Scaling
    type DAxisId = DocumentFormat.OpenXml.Drawing.Charts.AxisId
    type DNumberReference = DocumentFormat.OpenXml.Drawing.Charts.NumberReference
    type DStringReference = DocumentFormat.OpenXml.Drawing.Charts.StringReference
    type DValues = DocumentFormat.OpenXml.Drawing.Charts.Values
    type DCategoryAxisData = DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData
    type DSeriesText = DocumentFormat.OpenXml.Drawing.Charts.SeriesText
    type DLegend = DocumentFormat.OpenXml.Drawing.Charts.Legend
    type DLegendPosition = DocumentFormat.OpenXml.Drawing.Charts.LegendPosition
    type DTitle = DocumentFormat.OpenXml.Drawing.Charts.Title
    type DChartText = DocumentFormat.OpenXml.Drawing.Charts.ChartText
    type DRichText = DocumentFormat.OpenXml.Drawing.Charts.RichText
    type DAutoTitleDeleted = DocumentFormat.OpenXml.Drawing.Charts.AutoTitleDeleted
    type DPlotVisibleOnly = DocumentFormat.OpenXml.Drawing.Charts.PlotVisibleOnly
    type DFormula = DocumentFormat.OpenXml.Drawing.Charts.Formula
    type DAxisPosition = DocumentFormat.OpenXml.Drawing.Charts.AxisPosition
    type DDelete = DocumentFormat.OpenXml.Drawing.Charts.Delete
    type DMajorTickMark = DocumentFormat.OpenXml.Drawing.Charts.MajorTickMark
    type DMinorTickMark = DocumentFormat.OpenXml.Drawing.Charts.MinorTickMark
    type DTickLabelPosition = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition
    type DCrossingAxis = DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis
    type DAxisOrientation = DocumentFormat.OpenXml.Drawing.Charts.Orientation
    type DGapWidth = DocumentFormat.OpenXml.Drawing.Charts.GapWidth
    type DBarDirection = DocumentFormat.OpenXml.Drawing.Charts.BarDirection
    type DBarGrouping = DocumentFormat.OpenXml.Drawing.Charts.BarGrouping
    type DVaryColors = DocumentFormat.OpenXml.Drawing.Charts.VaryColors
    type DGrouping = DocumentFormat.OpenXml.Drawing.Charts.Grouping
    type DIndex = DocumentFormat.OpenXml.Drawing.Charts.Index
    type DOrder = DocumentFormat.OpenXml.Drawing.Charts.Order
    type DMajorGridlines = DocumentFormat.OpenXml.Drawing.Charts.MajorGridlines

    type DAxisPositionValues = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues
    type DTickMarkValues = DocumentFormat.OpenXml.Drawing.Charts.TickMarkValues
    type DTickLabelPositionValues = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues
    type DAxisOrientationValues = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues
    type DLegendPositionValues = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues
    type DBarDirectionValues = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues
    type DBarGroupingValues = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues
    type DGroupingValues = DocumentFormat.OpenXml.Drawing.Charts.GroupingValues

    type XChartReference = DocumentFormat.OpenXml.Drawing.Charts.ChartReference

    type XTwoCellAnchor = DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor
    type XFromMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker
    type XToMarker = DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker
    type XColumnId = DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId
    type XRowId = DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId
    type XColumnOffset = DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset
    type XRowOffset = DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset
    type XGraphicFrame = DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame
    type XNonVisualGraphicFrameProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties
    type XNonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties
    type XNonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
    type XTransform = DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform
    type XWorksheetDrawing = DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing
    type XClientData = DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData

    type AGraphic = DocumentFormat.OpenXml.Drawing.Graphic
    type AGraphicData = DocumentFormat.OpenXml.Drawing.GraphicData
    type AOffset = DocumentFormat.OpenXml.Drawing.Offset
    type AExtents = DocumentFormat.OpenXml.Drawing.Extents
    type ARun = DocumentFormat.OpenXml.Drawing.Run
    type AParagraph = DocumentFormat.OpenXml.Drawing.Paragraph
    type AText = DocumentFormat.OpenXml.Drawing.Text
    type ABodyProperties = DocumentFormat.OpenXml.Drawing.BodyProperties
    type AListStyle = DocumentFormat.OpenXml.Drawing.ListStyle

    // Only need to be unique within one chart's own <c:chart> (each chart is its own,
    // independent part) - not across the workbook, unlike `Table`'s ids.
    let private categoryAxisId = 111111111u
    let private valueAxisId = 222222222u

    let private quoteSheetName (name: string) = sprintf "'%s'" (name.Replace("'", "''"))
    let private absoluteCell (r: CellRef) = sprintf "$%s$%d" (CellRef.columnLetters r.Col) (r.Row + 1)

    let private chartCellReference (sheetName: string) (cell: CellRef) : string =
        sprintf "%s!%s" (quoteSheetName sheetName) (absoluteCell cell)

    let private chartRangeReference (sheetName: string) (topLeft: CellRef) (bottomRight: CellRef) : string =
        sprintf "%s!%s:%s" (quoteSheetName sheetName) (absoluteCell topLeft) (absoluteCell bottomRight)

    let private seriesTextElement (sheetName: string) (nameCell: CellRef) : DSeriesText =
        let sr = DStringReference()
        sr.Formula <- DFormula(chartCellReference sheetName nameCell)
        DSeriesText(StringReference = sr)

    let private categoryAxisDataElement (sheetName: string) (topLeft: CellRef) (bottomRight: CellRef) : DCategoryAxisData =
        let sr = DStringReference()
        sr.Formula <- DFormula(chartRangeReference sheetName topLeft bottomRight)
        DCategoryAxisData(StringReference = sr)

    let private valuesElement (sheetName: string) (topLeft: CellRef) (bottomRight: CellRef) : DValues =
        let nr = DNumberReference()
        nr.Formula <- DFormula(chartRangeReference sheetName topLeft bottomRight)
        DValues(NumberReference = nr)

    let private barSeriesElement
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (index: uint32)
        (series: ChartSeries)
        : DBarChartSeries =
        let ser = DBarChartSeries(Index = DIndex(Val = UInt32Value(index)), Order = DOrder(Val = UInt32Value(index)))
        ser.SeriesText <- seriesTextElement sheetName series.Name
        ser.AppendChild(categoryAxisDataElement sheetName categoriesTopLeft categoriesBottomRight) |> ignore
        ser.AppendChild(valuesElement sheetName series.ValuesTopLeft series.ValuesBottomRight) |> ignore
        ser

    let private lineSeriesElement
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (index: uint32)
        (series: ChartSeries)
        : DLineChartSeries =
        let ser = DLineChartSeries(Index = DIndex(Val = UInt32Value(index)), Order = DOrder(Val = UInt32Value(index)))
        ser.SeriesText <- seriesTextElement sheetName series.Name
        ser.AppendChild(categoryAxisDataElement sheetName categoriesTopLeft categoriesBottomRight) |> ignore
        ser.AppendChild(valuesElement sheetName series.ValuesTopLeft series.ValuesBottomRight) |> ignore
        ser

    let private pieSeriesElement
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (index: uint32)
        (series: ChartSeries)
        : DPieChartSeries =
        let ser = DPieChartSeries(Index = DIndex(Val = UInt32Value(index)), Order = DOrder(Val = UInt32Value(index)))
        ser.SeriesText <- seriesTextElement sheetName series.Name
        ser.AppendChild(categoryAxisDataElement sheetName categoriesTopLeft categoriesBottomRight) |> ignore
        ser.AppendChild(valuesElement sheetName series.ValuesTopLeft series.ValuesBottomRight) |> ignore
        ser

    let private barChartElement
        (direction: DBarDirectionValues)
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (series: ChartSeries list)
        : DBarChart =
        let chart =
            DBarChart(
                BarDirection = DBarDirection(Val = EnumValue<DBarDirectionValues>(direction)),
                BarGrouping = DBarGrouping(Val = EnumValue<DBarGroupingValues>(DBarGroupingValues.Clustered))
            )

        series
        |> List.iteri (fun i s -> chart.AppendChild(barSeriesElement sheetName categoriesTopLeft categoriesBottomRight (uint32 i) s) |> ignore)

        chart.AppendChild(DGapWidth(Val = UInt16Value(150us))) |> ignore
        chart.AppendChild(DAxisId(Val = UInt32Value(categoryAxisId))) |> ignore
        chart.AppendChild(DAxisId(Val = UInt32Value(valueAxisId))) |> ignore
        chart

    let private lineChartElement
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (series: ChartSeries list)
        : DLineChart =
        let chart = DLineChart(Grouping = DGrouping(Val = EnumValue<DGroupingValues>(DGroupingValues.Standard)))

        series
        |> List.iteri (fun i s -> chart.AppendChild(lineSeriesElement sheetName categoriesTopLeft categoriesBottomRight (uint32 i) s) |> ignore)

        chart.AppendChild(DAxisId(Val = UInt32Value(categoryAxisId))) |> ignore
        chart.AppendChild(DAxisId(Val = UInt32Value(valueAxisId))) |> ignore
        chart

    let private pieChartElement
        (sheetName: string)
        (categoriesTopLeft: CellRef)
        (categoriesBottomRight: CellRef)
        (series: ChartSeries list)
        : DPieChart =
        let chart = DPieChart(VaryColors = DVaryColors(Val = BooleanValue(true)))

        series
        |> List.iteri (fun i s -> chart.AppendChild(pieSeriesElement sheetName categoriesTopLeft categoriesBottomRight (uint32 i) s) |> ignore)

        chart

    let private axisScaling () : DScaling =
        DScaling(Orientation = DAxisOrientation(Val = EnumValue<DAxisOrientationValues>(DAxisOrientationValues.MinMax)))

    /// Builds the category axis. `position` differs between `ChartColumn`/`ChartLine`
    /// (bottom) and `ChartBar` (left, since a horizontal bar chart's category axis runs
    /// down the left side) - callers pass in whichever matches their chart type.
    let private categoryAxisElement (position: DAxisPositionValues) : DCategoryAxis =
        DCategoryAxis(
            AxisId = DAxisId(Val = UInt32Value(categoryAxisId)),
            Scaling = axisScaling (),
            Delete = DDelete(Val = BooleanValue(false)),
            AxisPosition = DAxisPosition(Val = EnumValue<DAxisPositionValues>(position)),
            MajorTickMark = DMajorTickMark(Val = EnumValue<DTickMarkValues>(DTickMarkValues.Outside)),
            MinorTickMark = DMinorTickMark(Val = EnumValue<DTickMarkValues>(DTickMarkValues.None)),
            TickLabelPosition = DTickLabelPosition(Val = EnumValue<DTickLabelPositionValues>(DTickLabelPositionValues.NextTo)),
            CrossingAxis = DCrossingAxis(Val = UInt32Value(valueAxisId))
        )

    /// Builds the value axis. `position` is the opposite of whatever the category axis
    /// used (see `categoryAxisElement`).
    let private valueAxisElement (position: DAxisPositionValues) : DValueAxis =
        DValueAxis(
            AxisId = DAxisId(Val = UInt32Value(valueAxisId)),
            Scaling = axisScaling (),
            Delete = DDelete(Val = BooleanValue(false)),
            AxisPosition = DAxisPosition(Val = EnumValue<DAxisPositionValues>(position)),
            MajorGridlines = DMajorGridlines(),
            MajorTickMark = DMajorTickMark(Val = EnumValue<DTickMarkValues>(DTickMarkValues.Outside)),
            MinorTickMark = DMinorTickMark(Val = EnumValue<DTickMarkValues>(DTickMarkValues.None)),
            TickLabelPosition = DTickLabelPosition(Val = EnumValue<DTickLabelPositionValues>(DTickLabelPositionValues.NextTo)),
            CrossingAxis = DCrossingAxis(Val = UInt32Value(categoryAxisId))
        )

    let private plotAreaElement (entry: ChartEntry) (sheetName: string) : DPlotArea =
        let plotArea = DPlotArea()
        let catTL, catBR = entry.CategoriesTopLeft, entry.CategoriesBottomRight

        match entry.Type with
        | ChartColumn ->
            plotArea.AppendChild(barChartElement DBarDirectionValues.Column sheetName catTL catBR entry.Series) |> ignore
            plotArea.AppendChild(categoryAxisElement DAxisPositionValues.Bottom) |> ignore
            plotArea.AppendChild(valueAxisElement DAxisPositionValues.Left) |> ignore
        | ChartBar ->
            plotArea.AppendChild(barChartElement DBarDirectionValues.Bar sheetName catTL catBR entry.Series) |> ignore
            plotArea.AppendChild(categoryAxisElement DAxisPositionValues.Left) |> ignore
            plotArea.AppendChild(valueAxisElement DAxisPositionValues.Bottom) |> ignore
        | ChartLine ->
            plotArea.AppendChild(lineChartElement sheetName catTL catBR entry.Series) |> ignore
            plotArea.AppendChild(categoryAxisElement DAxisPositionValues.Bottom) |> ignore
            plotArea.AppendChild(valueAxisElement DAxisPositionValues.Left) |> ignore
        | ChartPie -> plotArea.AppendChild(pieChartElement sheetName catTL catBR entry.Series) |> ignore

        plotArea

    /// A chart title is written as rich text (`c:rich`), the same shape real Excel uses,
    /// rather than a string-literal cache - simpler than reproducing the point-count/index
    /// bookkeeping `c:strLit` needs for what both render as identically in Excel.
    let private richTextTitle (text: string) : DChartText =
        let run = ARun(Text = AText(text))
        let paragraph = AParagraph()
        paragraph.AppendChild(run) |> ignore
        let richText = DRichText(BodyProperties = ABodyProperties(), ListStyle = AListStyle())
        richText.AppendChild(paragraph) |> ignore
        DChartText(RichText = richText)

    let private chartElement (entry: ChartEntry) (sheetName: string) : DChart =
        let chart = DChart()

        match entry.Title with
        | Some text -> chart.Title <- DTitle(ChartText = richTextTitle text)
        | None -> chart.AutoTitleDeleted <- DAutoTitleDeleted(Val = BooleanValue(true))

        chart.PlotArea <- plotAreaElement entry sheetName

        if entry.ShowLegend then
            chart.Legend <- DLegend(LegendPosition = DLegendPosition(Val = EnumValue<DLegendPositionValues>(DLegendPositionValues.Bottom)))

        chart.PlotVisibleOnly <- DPlotVisibleOnly(Val = BooleanValue(true))
        chart

    let private chartSpaceElement (entry: ChartEntry) (sheetName: string) : DChartSpace =
        let chartSpace = DChartSpace()
        chartSpace.AppendChild(chartElement entry sheetName) |> ignore
        chartSpace

    let private fromMarker (cell: CellRef) : XFromMarker =
        let m = XFromMarker()
        m.ColumnId <- XColumnId(string cell.Col)
        m.ColumnOffset <- XColumnOffset("0")
        m.RowId <- XRowId(string cell.Row)
        m.RowOffset <- XRowOffset("0")
        m

    /// One past `cell` (zero offset) - the "to" marker is the far edge of the anchor
    /// range, so a chart meant to span through the end of `BottomRightAnchor` anchors its
    /// "to" corner at the start of the *next* cell, not `BottomRightAnchor` itself.
    let private toMarker (cell: CellRef) : XToMarker =
        let m = XToMarker()
        m.ColumnId <- XColumnId(string (cell.Col + 1))
        m.ColumnOffset <- XColumnOffset("0")
        m.RowId <- XRowId(string (cell.Row + 1))
        m.RowOffset <- XRowOffset("0")
        m

    let private graphicFrameElement (chartId: uint32) (relId: string) : XGraphicFrame =
        let nvProps = XNonVisualGraphicFrameProperties()
        nvProps.NonVisualDrawingProperties <- XNonVisualDrawingProperties(Id = UInt32Value(chartId), Name = StringValue(sprintf "Chart %d" chartId))
        nvProps.NonVisualGraphicFrameDrawingProperties <- XNonVisualGraphicFrameDrawingProperties()

        let graphicData = AGraphicData(Uri = StringValue("http://schemas.openxmlformats.org/drawingml/2006/chart"))
        graphicData.AppendChild(XChartReference(Id = StringValue(relId))) |> ignore

        let frame = XGraphicFrame()
        frame.NonVisualGraphicFrameProperties <- nvProps
        frame.Transform <- XTransform(Offset = AOffset(X = Int64Value(0L), Y = Int64Value(0L)), Extents = AExtents(Cx = Int64Value(0L), Cy = Int64Value(0L)))
        frame.Graphic <- AGraphic(GraphicData = graphicData)
        frame

    let private twoCellAnchorElement (entry: ChartEntry) (chartId: uint32) (relId: string) : XTwoCellAnchor =
        let anchor = XTwoCellAnchor()
        anchor.FromMarker <- fromMarker entry.TopLeftAnchor
        anchor.ToMarker <- toMarker entry.BottomRightAnchor
        anchor.AppendChild(graphicFrameElement chartId relId) |> ignore
        anchor.AppendChild(XClientData()) |> ignore
        anchor

    /// Adds one `ChartPart` per chart to `drawingsPart` (already created by the caller -
    /// shared with `ImageWriter` when a sheet has both charts and images anchored on the
    /// same drawing canvas) and returns the anchor element for each, starting numbering
    /// at `startId` - object ids must be unique across every drawing object on the sheet,
    /// not just among charts, so the caller is responsible for reserving a disjoint id
    /// range per kind of object.
    let chartAnchors (drawingsPart: DrawingsPart) (sheetName: string) (startId: uint32) (charts: ChartEntry list) : OpenXmlElement list =
        charts
        |> List.mapi (fun i entry ->
            let chartPart = drawingsPart.AddNewPart<ChartPart>()
            chartPart.ChartSpace <- chartSpaceElement entry sheetName
            chartPart.ChartSpace.Save()

            let relId = drawingsPart.GetIdOfPart(chartPart)
            let chartId = startId + uint32 i
            twoCellAnchorElement entry chartId relId :> OpenXmlElement)
