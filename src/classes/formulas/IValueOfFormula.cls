VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueOfFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Surface contract for VALUE_OF formula parsers"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Surface contract for VALUE_OF formula parsers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IValueOfFormula
'@description
'Defines the public contract for the ValueOfFormula class, exposing
'validation, converted formula output, failure diagnostics, and lookup
'metadata for VALUE_OF custom expressions. Implementations validate
'VALUE_OF expressions from the dictionary and resolve sheet names and
'column indices for downstream consumers.


'@section Conversion
'===============================================================================
'@description Converted formula output and validation status.

'@jump:converted-formula
'@prop-title Retrieve the converted VALUE_OF expression ready for workbook usage
'@details
'Returns the string representation with sheet name and column indices
'resolved. Returns vbNullString when the formula is invalid.
'@return String. Converted VALUE_OF expression.
Public Property Get ConvertedFormula() As String: End Property

'@jump:valid
'@prop-title Report whether the VALUE_OF formula parsed successfully
'@details
'Implementations perform syntax checks and dictionary lookups before
'returning True. Returns False when any validation step fails.
'@return Boolean. True when all checks passed.
Public Property Get Valid() As Boolean: End Property

'@jump:failure-reason
'@prop-title Provide the human-readable reason describing the failure
'@details
'Returns a diagnostic message explaining why parsing or validation
'failed. Returns vbNullString when the formula is valid.
'@return String. Explanation of the failure.
Public Property Get FailureReason() As String: End Property


'@section LookupMetadata
'===============================================================================
'@description Resolved sheet name and column indices for lookup operations.

'@jump:lookup-sheet-name
'@prop-title Expose the resolved worksheet hosting the lookup table
'@details
'Returns the name of the worksheet containing the variable referenced
'by the VALUE_OF expression. Returns vbNullString when unavailable.
'@return String. Sheet name.
Public Property Get LookupSheetName() As String: End Property

'@jump:lookup-column-index
'@prop-title Return the column index used for the lookup key
'@details
'Indicates the column position within the lookup table that holds the
'key values. Returns zero when not resolved.
'@return Long. Column index.
Public Property Get LookupColumnIndex() As Long: End Property

'@jump:value-column-index
'@prop-title Return the column index providing the matched value
'@details
'Indicates the column position within the lookup table that holds the
'result values. Returns zero when not resolved.
'@return Long. Column index.
Public Property Get ValueColumnIndex() As Long: End Property
