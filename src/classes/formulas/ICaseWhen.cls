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
'@IgnoreModule SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  ProcedureNotUsed, UseMeaningfulName, HungarianNotation
'@Folder("Custom Functions")
'@ModuleDescription("Interface of CaseWhen")

Option Explicit


'@jump:Create
'@section Overview
'@details Describes the surface exposed by CaseWhen parsers: validity checks, nested IF output, and category extraction.


'@jump:Valid
'@sub-title Determine whether the supplied formula can be parsed
'@details Implementations should verify the presence of the CASE_WHEN token before attempting to parse.
'@return Boolean True when the formula is parseable.
Public Property Get Valid() As Boolean: End Property


'@jump:ParsedFormula
'@sub-title Retrieve the CASE_WHEN formula rewritten as nested IF statements
'@details Consumers can feed the resulting expression directly into worksheet formulas.
'@return String nested IF expression.
Public Property Get parsedFormula() As String: End Property


'@jump:Categories
'@sub-title Enumerate the textual results referenced by the CASE_WHEN expression
'@details Each element represents a display label corresponding to a formula branch.
'@return BetterArray containing quoted labels.
Public Property Get Categories() As BetterArray: End Property
