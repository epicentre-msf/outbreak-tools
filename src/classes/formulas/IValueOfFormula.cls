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

'@section Overview
'===============================================================================
'@details Implementations validate VALUE_OF expressions from the dictionary and expose the converted formula alongside lookup metadata for downstream consumers.

'@jump:ConvertedFormula
'@sub-title Retrieve the converted VALUE_OF expression ready for workbook usage
'@details Returns the string representation with sheet name and column indices resolved.
'@return String converted VALUE_OF expression, or vbNullString when invalid.
Public Property Get ConvertedFormula() As String: End Property

'@jump:Valid
'@sub-title Report whether the VALUE_OF formula parsed successfully
'@details Implementations perform syntax and dictionary lookups before returning True.
'@return Boolean True when all checks passed.
Public Property Get Valid() As Boolean: End Property

'@jump:FailureReason
'@sub-title Provide the human readable reason describing the failure
'@return String explanation; vbNullString when the formula is valid.
Public Property Get FailureReason() As String: End Property

'@jump:LookupSheetName
'@sub-title Expose the resolved worksheet hosting the lookup table
'@return String sheet name; vbNullString when unavailable.
Public Property Get LookupSheetName() As String: End Property

'@jump:LookupColumnIndex
'@sub-title Return the column index used for the lookup key
'@return Long column index; zero when not resolved.
Public Property Get LookupColumnIndex() As Long: End Property

'@jump:ValueColumnIndex
'@sub-title Return the column index providing the matched value
'@return Long column index; zero when not resolved.
Public Property Get ValueColumnIndex() As Long: End Property
