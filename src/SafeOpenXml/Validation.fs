namespace SafeOpenXml

/// Shared by both conditional formatting's `CellValueRule` and data validation's
/// numeric/text-length rules - OOXML uses the same comparison vocabulary for both
/// (`cellIs`'s `operator` attribute and `dataValidation`'s `operator` attribute).
type ComparisonOperator =
    | Equal
    | NotEqual
    | GreaterThan
    | LessThan
    | GreaterThanOrEqual
    | LessThanOrEqual
    /// Needs two operands (`formula1`/`formula2`).
    | Between
    /// Needs two operands (`formula1`/`formula2`).
    | NotBetween

/// A conditional formatting rule to apply over a range. `formula1`/`formula2` are raw
/// formula text (same convention as `CellValue.Formula`) - for `CellValueRule` these are
/// literal values or cell references compared against, not `=`-prefixed formulas.
///
/// Covers the common cases; icon sets, "top/bottom N", and the text/blank/error-contains
/// rule kinds aren't modeled - see MAPPING.md. Named `ColorScale2`/`ColorScale3`/
/// `DataBarRule` rather than bare `ColorScale`/`DataBar` because
/// `DocumentFormat.OpenXml.Spreadsheet` already defines types with those exact names,
/// same reasoning as `FontStyle`/`FillStyle`/`BorderStyle` elsewhere in this file's family.
type ConditionalFormatRule =
    | CellValueRule of operator: ComparisonOperator * formula1: string * formula2: string option * style: CellStyle
    | FormulaRule of formula: string * style: CellStyle
    | ColorScale2 of minColor: Color * maxColor: Color
    | ColorScale3 of minColor: Color * midColor: Color * maxColor: Color
    /// A single-color data bar with Excel's default automatic min/max thresholds.
    | DataBarRule of color: Color
    | DuplicateValuesRule of style: CellStyle
    | UniqueValuesRule of style: CellStyle

/// What kind of value a data validation rule accepts. `formula1`/`formula2` are raw
/// formula text, same convention as `ConditionalFormatRule`.
///
/// Covers the common cases; `Date`/`Time` validation and cross-sheet named-range list
/// sources aren't modeled - see MAPPING.md.
type ValidationKind =
    /// A fixed, inline dropdown list of choices.
    | ListValidation of items: string list
    /// A dropdown list sourced from another range's values.
    | ListFromRangeValidation of topLeft: CellRef * bottomRight: CellRef
    | WholeNumberValidation of operator: ComparisonOperator * formula1: string * formula2: string option
    | DecimalValidation of operator: ComparisonOperator * formula1: string * formula2: string option
    | TextLengthValidation of operator: ComparisonOperator * formula1: string * formula2: string option
    /// An arbitrary boolean formula.
    | CustomValidation of formula: string

type ErrorAlertStyle =
    | Stop
    | Warning
    | Information

/// The non-essential parts of a data validation: whether blanks are allowed, and the
/// optional input prompt / error alert shown to the user. Kept separate from
/// `ValidationKind` so the common case (just a rule, no custom messages) doesn't need to
/// mention any of this - see `SheetDsl.dataValidation`'s optional parameters.
type ValidationAlert =
    { AllowBlank: bool
      ErrorStyle: ErrorAlertStyle
      ErrorTitle: string option
      ErrorMessage: string option
      InputTitle: string option
      InputMessage: string option }

    static member Default =
        { AllowBlank = true
          ErrorStyle = Stop
          ErrorTitle = None
          ErrorMessage = None
          InputTitle = None
          InputMessage = None }

/// A conditional formatting rule applied over a range, as it's stored on `Worksheet`.
type ConditionalFormatEntry =
    { TopLeft: CellRef
      BottomRight: CellRef
      Rule: ConditionalFormatRule }

/// A data validation rule applied over a range, as it's stored on `Worksheet`.
type DataValidationEntry =
    { TopLeft: CellRef
      BottomRight: CellRef
      Kind: ValidationKind
      Alert: ValidationAlert }
