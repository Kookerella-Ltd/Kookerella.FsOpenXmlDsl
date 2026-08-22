namespace SafeOpenXml.Interpreter

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open SafeOpenXml

/// Parses the pivot tables anchored on one worksheet back into `PivotTableEntry` values -
/// the read side of `PivotTableWriter`. Best-effort in the same sense as the rest of this
/// codebase, but more narrowly so: a pivot table this library didn't write itself can use
/// far more of the format than `PivotTableEntry` models (subtotals, multiple/nested row
/// or column fields, page filters, ...), so this only round-trips pivot tables matching
/// the one shape `PivotTableWriter` produces (exactly one row field, at most one column
/// field, exactly one value field) - anything else is dropped rather than guessed at.
module internal PivotTableReader =

    let private aggregationOf (v: DataConsolidateFunctionValues) : PivotAggregation option =
        if v = DataConsolidateFunctionValues.Sum then Some PivotSum
        elif v = DataConsolidateFunctionValues.Count then Some PivotCount
        elif v = DataConsolidateFunctionValues.CountNumbers then Some PivotCountNumbers
        elif v = DataConsolidateFunctionValues.Average then Some PivotAverage
        elif v = DataConsolidateFunctionValues.Minimum then Some PivotMin
        elif v = DataConsolidateFunctionValues.Maximum then Some PivotMax
        else None

    let private aggregationLabel (agg: PivotAggregation) : string =
        match agg with
        | PivotSum -> "Sum"
        | PivotCount -> "Count"
        | PivotCountNumbers -> "Count Numbers"
        | PivotAverage -> "Average"
        | PivotMin -> "Min"
        | PivotMax -> "Max"

    let private parseRange (reference: string) : CellRef * CellRef =
        let parts = reference.Split(':')
        let topLeft = CellRef.ofA1 parts.[0]
        let bottomRight = CellRef.ofA1 (if parts.Length > 1 then parts.[1] else parts.[0])
        (topLeft, bottomRight)

    /// Tries to interpret one `PivotTablePart` as a `PivotTableEntry` - `None` if it
    /// doesn't match the single-row-field/at-most-one-column-field/single-value-field
    /// shape `PivotTableWriter` always produces.
    let private tryReadOne (destinationSheetName: string) (pivotTablePart: PivotTablePart) : PivotTableEntry option =
        let definition = pivotTablePart.PivotTableDefinition
        let cacheDefPart = pivotTablePart.PivotTableCacheDefinitionPart

        if isNull definition || isNull cacheDefPart then
            None
        else
            let cacheDefinition = cacheDefPart.PivotCacheDefinition

            let fieldNames =
                match Option.ofObj cacheDefinition |> Option.bind (fun cd -> Option.ofObj cd.CacheFields) with
                | Some cacheFields -> cacheFields.Elements<CacheField>() |> Seq.map (fun f -> f.Name.Value) |> Array.ofSeq
                | None -> [||]

            let nameAt (i: int) : string option = if i >= 0 && i < fieldNames.Length then Some fieldNames.[i] else None

            let rowFieldName =
                Option.ofObj definition.RowFields
                |> Option.bind (fun rf -> rf.Elements<Field>() |> Seq.tryHead)
                |> Option.bind (fun f -> Option.ofObj f.Index)
                |> Option.bind (fun idx -> nameAt idx.Value)

            let onlyOneRowField =
                match Option.ofObj definition.RowFields with
                | Some rf -> rf.Elements<Field>() |> Seq.length = 1
                | None -> false

            let columnFieldName =
                Option.ofObj definition.ColumnFields
                |> Option.bind (fun cf -> cf.Elements<Field>() |> Seq.tryHead)
                |> Option.bind (fun f -> Option.ofObj f.Index)
                |> Option.bind (fun idx -> nameAt idx.Value)

            let atMostOneColumnField =
                match Option.ofObj definition.ColumnFields with
                | None -> true
                | Some cf -> cf.Elements<Field>() |> Seq.length = 1

            let dataFieldInfo =
                Option.ofObj definition.DataFields
                |> Option.bind (fun df -> df.Elements<DataField>() |> Seq.tryHead)
                |> Option.bind (fun df ->
                    match Option.ofObj df.Field, Option.ofObj df.Subtotal |> Option.bind (fun s -> aggregationOf s.Value) with
                    | Some fld, Some agg -> nameAt (int fld.Value) |> Option.map (fun name -> name, agg, (if isNull df.Name then None else Some df.Name.Value))
                    | _ -> None)

            let onlyOneDataField =
                match Option.ofObj definition.DataFields with
                | Some df -> df.Elements<DataField>() |> Seq.length = 1
                | None -> false

            let sourceInfo =
                Option.ofObj cacheDefinition
                |> Option.bind (fun cd -> Option.ofObj cd.CacheSource)
                |> Option.bind (fun cs -> Option.ofObj cs.WorksheetSource)
                |> Option.bind (fun ws ->
                    match Option.ofObj ws.Reference, Option.ofObj ws.Sheet with
                    | Some refVal, Some sheetVal -> Some(refVal.Value, sheetVal.Value)
                    | _ -> None)

            let anchor =
                Option.ofObj definition.Location |> Option.bind (fun loc -> Option.ofObj loc.Reference) |> Option.map (fun r -> fst (parseRange r.Value))

            match rowFieldName, dataFieldInfo, sourceInfo, anchor with
            | Some rowField, Some(valueField, agg, dataFieldCaption), Some(sourceRef, sourceSheetName), Some topLeftAnchor when
                onlyOneRowField && atMostOneColumnField && onlyOneDataField ->
                let sourceTopLeft, sourceBottomRight = parseRange sourceRef
                let defaultCaption = sprintf "%s of %s" (aggregationLabel agg) valueField

                let valueCaption =
                    match dataFieldCaption with
                    | Some c when c <> defaultCaption -> Some c
                    | _ -> None

                Some
                    { SourceSheet = if sourceSheetName = destinationSheetName then None else Some sourceSheetName
                      SourceTopLeft = sourceTopLeft
                      SourceBottomRight = sourceBottomRight
                      RowField = rowField
                      ColumnField = columnFieldName
                      ValueField = valueField
                      Aggregation = agg
                      ValueCaption = valueCaption
                      TopLeftAnchor = topLeftAnchor }
            | _ -> None

    let readPivotTables (worksheetPart: WorksheetPart) (destinationSheetName: string) : PivotTableEntry list =
        worksheetPart.PivotTableParts |> Seq.choose (tryReadOne destinationSheetName) |> List.ofSeq
