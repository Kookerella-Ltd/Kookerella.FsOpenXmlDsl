namespace Kookerella.FsOpenXmlDsl.Interpreter

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open Kookerella.FsOpenXmlDsl

/// Builds the pivot cache and pivot table definition parts for the pivot tables anchored
/// on one worksheet - the write side of `PivotTableReader`. Unlike every other feature in
/// this library, this one isn't a pure translation: `computeGrid` actually performs the
/// grouping and aggregation `PivotTableEntry` describes, since a pivot table's file
/// format bakes in the *result* of that computation in three places (the cache records,
/// the row/column item layout, and the literal cell values written into the worksheet)
/// that all have to agree - see `PivotTableEntry`'s own doc comment.
///
/// None of the pivot-related types here collide with anything already qualified
/// elsewhere in this codebase (unlike DrawingML/ChartML), so `DocumentFormat.OpenXml.
/// Spreadsheet` is opened directly rather than needing type abbreviations - see
/// `EmbeddedPivotTable`'s own doc comment in `Builders.fs`.
module internal PivotTableWriter =

    // --- The aggregation engine -----------------------------------------------------

    /// Reads the source range's header row (must be `Text` cells) and data rows from
    /// whichever sheet holds them - `allSheets` is the whole workbook's sheets, since a
    /// pivot table's source range is very often on a different sheet than the pivot table
    /// itself.
    let private sourceData (allSheets: Worksheet list) (destinationSheetName: string) (entry: PivotTableEntry) : string[] * CellValue[][] =
        let sourceSheetName = entry.SourceSheet |> Option.defaultValue destinationSheetName

        let sourceSheet =
            match allSheets |> List.tryFind (fun s -> s.Name = sourceSheetName) with
            | Some s -> s
            | None -> invalidArg (nameof entry) (sprintf "Pivot table's source sheet '%s' isn't in this workbook" sourceSheetName)

        let cellMap = sourceSheet.Cells |> List.map (fun c -> c.Ref, c.Value) |> Map.ofList
        let headerRow = entry.SourceTopLeft.Row

        let headers =
            [| for c in entry.SourceTopLeft.Col .. entry.SourceBottomRight.Col ->
                   match cellMap.TryFind(CellRef.create headerRow c) with
                   | Some(Text s) -> s
                   | _ ->
                       invalidArg
                           (nameof entry)
                           (sprintf "Pivot table source header at %s must be a text cell" (CellRef.toA1 (CellRef.create headerRow c))) |]

        let dataRows =
            [| for r in headerRow + 1 .. entry.SourceBottomRight.Row ->
                   [| for c in entry.SourceTopLeft.Col .. entry.SourceBottomRight.Col ->
                          cellMap.TryFind(CellRef.create r c) |> Option.defaultValue Empty |] |]

        (headers, dataRows)

    let private fieldIndex (headers: string[]) (entry: PivotTableEntry) (fieldName: string) : int =
        match Array.tryFindIndex ((=) fieldName) headers with
        | Some i -> i
        | None -> invalidArg (nameof entry) (sprintf "Pivot table field '%s' isn't a header in the source range" fieldName)

    let private numericValueOf (v: CellValue) : float option =
        match v with
        | Number n -> Some n
        | Formula(_, Some cached) -> Some cached
        | Boolean b -> Some(if b then 1.0 else 0.0)
        | _ -> None

    let private aggregationLabel (agg: PivotAggregation) : string =
        match agg with
        | PivotSum -> "Sum"
        | PivotCount -> "Count"
        | PivotCountNumbers -> "Count Numbers"
        | PivotAverage -> "Average"
        | PivotMin -> "Min"
        | PivotMax -> "Max"

    let private aggregate (agg: PivotAggregation) (values: CellValue list) : float =
        match agg with
        | PivotCount -> values |> List.filter (fun v -> v <> Empty) |> List.length |> float
        | PivotCountNumbers -> values |> List.choose numericValueOf |> List.length |> float
        | PivotSum -> values |> List.choose numericValueOf |> List.sum
        | PivotAverage ->
            match values |> List.choose numericValueOf with
            | [] -> 0.0
            | nums -> List.sum nums / float nums.Length
        | PivotMin ->
            match values |> List.choose numericValueOf with
            | [] -> 0.0
            | nums -> List.min nums
        | PivotMax ->
            match values |> List.choose numericValueOf with
            | [] -> 0.0
            | nums -> List.max nums

    /// A stable, total ordering over `CellValue` for sorting distinct field values -
    /// numbers compare numerically, text ordinally, and different types fall back to a
    /// fixed rank so a mixed-type column never throws, it just sorts less meaningfully.
    let private compareCellValue (a: CellValue) (b: CellValue) : int =
        let rank =
            function
            | Empty -> 0
            | Number _ -> 1
            | Date _ -> 2
            | Boolean _ -> 3
            | Text _ -> 4
            | Formula _ -> 5

        match a, b with
        | Number x, Number y -> compare x y
        | Text x, Text y -> System.String.CompareOrdinal(x, y)
        | Date x, Date y -> compare x y
        | Boolean x, Boolean y -> compare x y
        | _ -> compare (rank a) (rank b)

    /// Computes the displayed pivot table grid - a Tabular-form layout (one column per
    /// row field/column value, no field indentation) with a grand total row and, when
    /// there's a column field, a grand total column too. Returns the cells to merge into
    /// the destination worksheet, plus the field indices and sorted distinct keys the
    /// caller needs to build the matching pivot table definition.
    let computeGrid
        (allSheets: Worksheet list)
        (destinationSheetName: string)
        (entry: PivotTableEntry)
        : {| Cells: Cell list
             RowFieldIndex: int
             ColumnFieldIndex: int option
             ValueFieldIndex: int
             RowKeys: CellValue[]
             ColumnKeys: CellValue[] option
             SourceHeaders: string[] |}
        =
        let headers, dataRows = sourceData allSheets destinationSheetName entry
        let rowFieldIdx = fieldIndex headers entry entry.RowField
        let colFieldIdx = entry.ColumnField |> Option.map (fieldIndex headers entry)
        let valueFieldIdx = fieldIndex headers entry entry.ValueField

        let rowKeys =
            dataRows |> Array.map (fun r -> r.[rowFieldIdx]) |> Array.distinct |> Array.sortWith compareCellValue

        let colKeysOpt =
            colFieldIdx
            |> Option.map (fun ci -> dataRows |> Array.map (fun r -> r.[ci]) |> Array.distinct |> Array.sortWith compareCellValue)

        let valuesWhere (predicate: CellValue[] -> bool) : CellValue list =
            dataRows |> Array.filter predicate |> Array.map (fun r -> r.[valueFieldIdx]) |> List.ofArray

        let caption =
            entry.ValueCaption
            |> Option.defaultValue (sprintf "%s of %s" (aggregationLabel entry.Aggregation) entry.ValueField)

        let anchor = entry.TopLeftAnchor
        let cellAt (r: int) (c: int) (v: CellValue) : Cell = { Ref = CellRef.create (anchor.Row + r) (anchor.Col + c); Value = v; Style = None }

        let cells =
            match colKeysOpt with
            | None ->
                let headerRow = [ cellAt 0 0 (Text entry.RowField); cellAt 0 1 (Text caption) ]

                let dataCellRows =
                    rowKeys
                    |> Array.toList
                    |> List.mapi (fun i rowKey ->
                        let agg = aggregate entry.Aggregation (valuesWhere (fun r -> r.[rowFieldIdx] = rowKey))
                        [ cellAt (i + 1) 0 rowKey; cellAt (i + 1) 1 (Number agg) ])

                let grandRowIdx = rowKeys.Length + 1
                let grandTotal = aggregate entry.Aggregation (valuesWhere (fun _ -> true))
                let grandRow = [ cellAt grandRowIdx 0 (Text "Grand Total"); cellAt grandRowIdx 1 (Number grandTotal) ]

                headerRow @ List.concat dataCellRows @ grandRow
            | Some colKeys ->
                let ci = colFieldIdx.Value

                let headerRow =
                    (cellAt 0 0 (Text entry.RowField))
                    :: (colKeys |> Array.toList |> List.mapi (fun j ck -> cellAt 0 (j + 1) ck))
                    @ [ cellAt 0 (colKeys.Length + 1) (Text "Grand Total") ]

                let dataCellRows =
                    rowKeys
                    |> Array.toList
                    |> List.mapi (fun i rowKey ->
                        let rowCells =
                            colKeys
                            |> Array.toList
                            |> List.mapi (fun j ck ->
                                let agg = aggregate entry.Aggregation (valuesWhere (fun r -> r.[rowFieldIdx] = rowKey && r.[ci] = ck))
                                cellAt (i + 1) (j + 1) (Number agg))

                        let rowGrandTotal = aggregate entry.Aggregation (valuesWhere (fun r -> r.[rowFieldIdx] = rowKey))
                        (cellAt (i + 1) 0 rowKey) :: rowCells @ [ cellAt (i + 1) (colKeys.Length + 1) (Number rowGrandTotal) ])

                let grandRowIdx = rowKeys.Length + 1

                let colGrandTotals =
                    colKeys
                    |> Array.toList
                    |> List.mapi (fun j ck ->
                        let agg = aggregate entry.Aggregation (valuesWhere (fun r -> r.[ci] = ck))
                        cellAt grandRowIdx (j + 1) (Number agg))

                let overallGrandTotal = aggregate entry.Aggregation (valuesWhere (fun _ -> true))

                let grandRow =
                    (cellAt grandRowIdx 0 (Text "Grand Total")) :: colGrandTotals @ [ cellAt grandRowIdx (colKeys.Length + 1) (Number overallGrandTotal) ]

                headerRow @ List.concat dataCellRows @ grandRow

        {| Cells = cells
           RowFieldIndex = rowFieldIdx
           ColumnFieldIndex = colFieldIdx
           ValueFieldIndex = valueFieldIdx
           RowKeys = rowKeys
           ColumnKeys = colKeysOpt
           SourceHeaders = headers |}

    // --- OOXML element construction ---------------------------------------------------

    let private sharedItemValueElement (v: CellValue) : OpenXmlElement =
        match v with
        | Text s -> StringItem(Val = StringValue(s))
        | Number n -> NumberItem(Val = DoubleValue(n))
        | Boolean b -> BooleanItem(Val = BooleanValue(b))
        | Date d -> DateTimeItem(Val = DateTimeValue(d))
        | Empty -> MissingItem()
        | Formula(_, Some cached) -> NumberItem(Val = DoubleValue(cached))
        | Formula(_, None) -> MissingItem()

    /// Builds one `cacheField` - `distinctValues` is `None` for a field this pivot table
    /// doesn't group by (the value field, and any other unused source column), which just
    /// gets a minimal `sharedItems` with no listed values.
    let private cacheFieldElement (name: string) (distinctValues: CellValue[] option) : CacheField =
        let cf = CacheField(Name = StringValue(name))

        let sharedItems =
            match distinctValues with
            | None -> SharedItems(ContainsSemiMixedTypes = BooleanValue(false), ContainsString = BooleanValue(false), ContainsNumber = BooleanValue(true))
            | Some values ->
                let containsString = values |> Array.exists (function Text _ -> true | _ -> false)
                let si = SharedItems(ContainsString = BooleanValue(containsString), ContainsBlank = BooleanValue(Array.contains Empty values))
                values |> Array.iter (fun v -> si.AppendChild(sharedItemValueElement v) |> ignore)
                si

        cf.SharedItems <- sharedItems
        cf

    let private pivotCacheDefinitionElement
        (sourceSheetName: string)
        (entry: PivotTableEntry)
        (grid:
            {| Cells: Cell list
               RowFieldIndex: int
               ColumnFieldIndex: int option
               ValueFieldIndex: int
               RowKeys: CellValue[]
               ColumnKeys: CellValue[] option
               SourceHeaders: string[] |})
        (recordCount: int)
        : PivotCacheDefinition =
        let sourceRange = sprintf "%s:%s" (CellRef.toA1 entry.SourceTopLeft) (CellRef.toA1 entry.SourceBottomRight)
        let worksheetSource = WorksheetSource(Reference = StringValue(sourceRange), Sheet = StringValue(sourceSheetName))
        let cacheSource = CacheSource(Type = EnumValue<SourceValues>(SourceValues.Worksheet))
        cacheSource.WorksheetSource <- worksheetSource

        let cacheFields = CacheFields(Count = UInt32Value(uint32 grid.SourceHeaders.Length))

        grid.SourceHeaders
        |> Array.iteri (fun i name ->
            let distinctValues =
                if i = grid.RowFieldIndex then Some grid.RowKeys
                elif Some i = grid.ColumnFieldIndex then grid.ColumnKeys
                else None

            cacheFields.AppendChild(cacheFieldElement name distinctValues) |> ignore)

        let def =
            PivotCacheDefinition(
                SaveData = BooleanValue(true),
                RefreshOnLoad = BooleanValue(true),
                RecordCount = UInt32Value(uint32 recordCount)
            )

        def.CacheSource <- cacheSource
        def.CacheFields <- cacheFields
        def

    let private pivotCacheRecordsElement (dataRows: CellValue[][]) : PivotCacheRecords =
        let records = PivotCacheRecords(Count = UInt32Value(uint32 dataRows.Length))

        dataRows
        |> Array.iter (fun row ->
            let record = PivotCacheRecord()
            row |> Array.iter (fun v -> record.AppendChild(sharedItemValueElement v) |> ignore)
            records.AppendChild(record) |> ignore)

        records

    let private dataConsolidateFunction (agg: PivotAggregation) : DataConsolidateFunctionValues =
        match agg with
        | PivotSum -> DataConsolidateFunctionValues.Sum
        | PivotCount -> DataConsolidateFunctionValues.Count
        | PivotCountNumbers -> DataConsolidateFunctionValues.CountNumbers
        | PivotAverage -> DataConsolidateFunctionValues.Average
        | PivotMin -> DataConsolidateFunctionValues.Minimum
        | PivotMax -> DataConsolidateFunctionValues.Maximum

    let private itemsElement (distinctCount: int) : Items =
        let items = Items()

        for i in 0 .. distinctCount - 1 do
            items.AppendChild(Item(Index = UInt32Value(uint32 i))) |> ignore

        let defaultItem = Item()
        defaultItem.ItemType <- EnumValue<ItemValues>(ItemValues.Default)
        items.AppendChild(defaultItem) |> ignore
        items

    let private pivotTableDefinitionElement
        (entry: PivotTableEntry)
        (cacheId: uint32)
        (grid:
            {| Cells: Cell list
               RowFieldIndex: int
               ColumnFieldIndex: int option
               ValueFieldIndex: int
               RowKeys: CellValue[]
               ColumnKeys: CellValue[] option
               SourceHeaders: string[] |})
        : PivotTableDefinition =
        let refs = grid.Cells |> List.map (fun c -> c.Ref)
        let minRow = refs |> List.map (fun r -> r.Row) |> List.min
        let maxRow = refs |> List.map (fun r -> r.Row) |> List.max
        let minCol = refs |> List.map (fun r -> r.Col) |> List.min
        let maxCol = refs |> List.map (fun r -> r.Col) |> List.max
        let wholeRangeReference = sprintf "%s:%s" (CellRef.toA1 (CellRef.create minRow minCol)) (CellRef.toA1 (CellRef.create maxRow maxCol))

        let location =
            Location(
                Reference = StringValue(wholeRangeReference),
                FirstHeaderRow = UInt32Value(0u),
                FirstDataRow = UInt32Value(1u),
                FirstDataColumn = UInt32Value(1u)
            )

        let pivotFields = PivotFields(Count = UInt32Value(uint32 grid.SourceHeaders.Length))

        grid.SourceHeaders
        |> Array.iteri (fun i _ ->
            let pf =
                if i = grid.RowFieldIndex then
                    let p = PivotField(Axis = EnumValue<PivotTableAxisValues>(PivotTableAxisValues.AxisRow), ShowAll = BooleanValue(false))
                    p.Items <- itemsElement grid.RowKeys.Length
                    p
                elif Some i = grid.ColumnFieldIndex then
                    let p = PivotField(Axis = EnumValue<PivotTableAxisValues>(PivotTableAxisValues.AxisColumn), ShowAll = BooleanValue(false))
                    p.Items <- itemsElement grid.ColumnKeys.Value.Length
                    p
                elif i = grid.ValueFieldIndex then
                    PivotField(DataField = BooleanValue(true), ShowAll = BooleanValue(false))
                else
                    PivotField(ShowAll = BooleanValue(false))

            pivotFields.AppendChild(pf) |> ignore)

        let rowFields = RowFields(Count = UInt32Value(1u))
        rowFields.AppendChild(Field(Index = Int32Value(grid.RowFieldIndex))) |> ignore

        let rowItems = RowItems()

        for i in 0 .. grid.RowKeys.Length - 1 do
            let item = RowItem()
            item.AppendChild(FieldItem(Val = UInt32Value(uint32 i))) |> ignore
            rowItems.AppendChild(item) |> ignore

        let grandRowItem = RowItem(ItemType = EnumValue<ItemValues>(ItemValues.Grand))
        rowItems.AppendChild(grandRowItem) |> ignore
        rowItems.Count <- UInt32Value(uint32 (grid.RowKeys.Length + 1))

        let dataFields = DataFields(Count = UInt32Value(1u))

        let dataFieldCaption =
            entry.ValueCaption
            |> Option.defaultValue (sprintf "%s of %s" (aggregationLabel entry.Aggregation) entry.ValueField)

        dataFields.AppendChild(
            DataField(Name = StringValue(dataFieldCaption), Field = UInt32Value(uint32 grid.ValueFieldIndex), Subtotal = EnumValue<DataConsolidateFunctionValues>(dataConsolidateFunction entry.Aggregation))
        )
        |> ignore

        let def =
            PivotTableDefinition(
                Name = StringValue("PivotTable1"),
                CacheId = UInt32Value(cacheId),
                DataCaption = StringValue("Values"),
                RowGrandTotals = BooleanValue(true),
                ColumnGrandTotals = BooleanValue(grid.ColumnFieldIndex.IsSome)
            )

        def.Location <- location
        def.PivotFields <- pivotFields
        def.RowFields <- rowFields
        def.RowItems <- rowItems
        def.DataFields <- dataFields

        match grid.ColumnFieldIndex with
        | Some colIdx ->
            let colFields = ColumnFields(Count = UInt32Value(1u))
            colFields.AppendChild(Field(Index = Int32Value(colIdx))) |> ignore
            def.ColumnFields <- colFields

            let colItems = ColumnItems()

            for i in 0 .. grid.ColumnKeys.Value.Length - 1 do
                let item = RowItem()
                item.AppendChild(FieldItem(Val = UInt32Value(uint32 i))) |> ignore
                colItems.AppendChild(item) |> ignore

            colItems.AppendChild(RowItem(ItemType = EnumValue<ItemValues>(ItemValues.Grand))) |> ignore
            colItems.Count <- UInt32Value(uint32 (grid.ColumnKeys.Value.Length + 1))
            def.ColumnItems <- colItems
        | None -> ()

        def.PivotTableStyle <- PivotTableStyle(Name = StringValue("PivotStyleLight16"), ShowRowHeaders = BooleanValue(true), ShowColumnHeaders = BooleanValue(true), ShowRowStripes = BooleanValue(false))

        def

    /// Builds the pivot cache and pivot table definition parts for one `PivotTableEntry`,
    /// and returns the workbook-level relationship id (from `workbookPart` to the new
    /// cache definition part) the caller needs to build a `<pivotCache>` entry in
    /// `workbook.xml`'s `<pivotCaches>` - see `Writer.populate`.
    let addPivotTable
        (workbookPart: WorkbookPart)
        (worksheetPart: WorksheetPart)
        (allSheets: Worksheet list)
        (destinationSheetName: string)
        (cacheId: uint32)
        (entry: PivotTableEntry)
        : string =
        let _, dataRows = sourceData allSheets destinationSheetName entry
        let sourceSheetName = entry.SourceSheet |> Option.defaultValue destinationSheetName
        let grid = computeGrid allSheets destinationSheetName entry

        let cacheDefPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>()
        cacheDefPart.PivotCacheDefinition <- pivotCacheDefinitionElement sourceSheetName entry grid dataRows.Length
        cacheDefPart.PivotCacheDefinition.Save()

        let recordsPart = cacheDefPart.AddNewPart<PivotTableCacheRecordsPart>()
        recordsPart.PivotCacheRecords <- pivotCacheRecordsElement dataRows
        recordsPart.PivotCacheRecords.Save()

        let pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>()
        pivotTablePart.PivotTableDefinition <- pivotTableDefinitionElement entry cacheId grid
        pivotTablePart.PivotTableDefinition.Save()
        pivotTablePart.AddPart(cacheDefPart) |> ignore

        workbookPart.GetIdOfPart(cacheDefPart)
