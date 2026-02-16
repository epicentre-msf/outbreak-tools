VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChoiceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for ChoiceFormula parsers"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Interface for ChoiceFormula parsers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IChoiceFormula
'@description
'Defines the public contract for the ChoiceFormula class, exposing
'validation, CASE_WHEN conversion, category extraction, and choice
'list identification for CHOICE_FORMULA custom expressions.


'@section Operations
'===============================================================================
'@description Validation, conversion, and metadata extraction.

'@jump:valid
'@prop-title Determine whether the stored expression is a valid CHOICE_FORMULA
'@details
'Implementations verify the CHOICE_FORMULA token, balanced parentheses,
'and that at least three arguments are present (name plus one
'condition/result pair). Returns False for blank arguments.
'@return Boolean. True when the specification can be parsed.
Public Property Get Valid() As Boolean: End Property

'@jump:parsed-formula
'@prop-title Retrieve the CHOICE_FORMULA expression rewritten as nested IF statements
'@details
'Delegates to the CaseWhen converter to produce the nested IF output.
'Returns an empty string when the formula is invalid.
'@return String. Nested IF expression.
Public Property Get parsedFormula() As String: End Property

'@jump:categories
'@prop-title Enumerate the textual labels referenced by the CHOICE_FORMULA
'@details
'Each element represents a display label shown to users for a given
'branch of the formula. Delegates category extraction to CaseWhen.
'@return BetterArray. Collection containing quoted labels.
Public Property Get Categories() As BetterArray: End Property

'@jump:choice-name
'@prop-title Retrieve the name of the referenced choice list
'@details
'Returns the first argument of the CHOICE_FORMULA expression, which
'identifies the choice list providing valid options.
'@return String. Identifier for the choice list.
Public Property Get choiceName() As String: End Property
