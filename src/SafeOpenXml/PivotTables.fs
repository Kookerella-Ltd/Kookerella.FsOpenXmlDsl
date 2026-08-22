namespace SafeOpenXml

/// How a pivot table's value field is aggregated across each group of matching source
/// rows. Covers the six most commonly used consolidation functions; see MAPPING.md for
/// the remainder (Product, StdDev, StdDevP, Var, VarP) that aren't modeled.
type PivotAggregation =
    | PivotSum
    | PivotCount
    | PivotCountNumbers
    | PivotAverage
    | PivotMin
    | PivotMax

/// A pivot table computed from a range of source cells and displayed as a grid anchored
/// at `TopLeftAnchor`, as stored on `Worksheet` (a sheet can have several). Unlike every
/// other feature in this library, this one isn't a pure translation of a DSL value into
/// OOXML - a pivot table's file format bakes in the *result* of an aggregation (grouped,
/// summarized source data) in three places that all have to agree (the pivot cache, the
/// pivot table definition's row/column layout, and the literal computed values written
/// into the worksheet grid itself), so `Writer` actually performs the grouping and
/// aggregation described here rather than just describing a reference for Excel to
/// resolve later.
///
/// Scoped to the single most common pivot table shape to keep that computation
/// tractable and verifiably correct: exactly one row field, at most one column field, and
/// exactly one value field - no nested row/column fields, no page/filter fields, no
/// subtotals beyond the grand total row/column. See MAPPING.md for what a real pivot
/// table can do that this doesn't model.
type PivotTableEntry =
    { /// `None` means the source range is on the same sheet as the pivot table itself.
      SourceSheet: string option
      SourceTopLeft: CellRef
      SourceBottomRight: CellRef
      /// Must exactly match a header cell's text in the source range's first row.
      RowField: string
      /// Must exactly match a header cell's text in the source range's first row, if set.
      ColumnField: string option
      /// Must exactly match a header cell's text in the source range's first row.
      ValueField: string
      Aggregation: PivotAggregation
      /// Column header for the aggregated value - defaults to Excel's own convention
      /// (e.g. `"Sum of Sales"`) when `None`.
      ValueCaption: string option
      TopLeftAnchor: CellRef }
