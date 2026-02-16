VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICaseWhen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of CaseWhen"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Interface of CaseWhen")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ICaseWhen
'@description
'Defines the public contract for the CaseWhen class, exposing validity
'checks, nested IF output, and category extraction for CASE_WHEN
'custom formula expressions.


'@section Operations
'===============================================================================
'@description Validation, conversion, and label extraction.

'@jump:valid
'@prop-title Determine whether the supplied formula can be parsed
'@details
'Implementations verify the presence of the CASE_WHEN( token at the
'start of the formula before attempting to parse. Returns False for
'empty or non-matching strings.
'@return Boolean. True when the formula is parseable.
Public Property Get Valid() As Boolean: End Property

'@jump:parsed-formula
'@prop-title Retrieve the CASE_WHEN formula rewritten as nested IF statements
'@details
'Consumers can feed the resulting expression directly into worksheet
'formulas. Returns an empty string when the formula is invalid.
'@return String. Nested IF expression.
Public Property Get parsedFormula() As String: End Property

'@jump:categories
'@prop-title Enumerate the textual results referenced by the CASE_WHEN expression
'@details
'Each element represents a display label corresponding to a formula
'branch. Quoted result entries are stripped; non-quoted entries are
'wrapped in quotes for consistent handling.
'@return BetterArray. Collection containing quoted labels.
Public Property Get Categories() As BetterArray: End Property
