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
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation
'@Folder("Custom Functions")
'@ModuleDescription("Interface for ChoiceFormula parsers")

Option Explicit

'@section Overview
'===============================================================================

'@details Describes the members exposed by ChoiceFormula: validation, CASE_WHEN conversion, category extraction, and choice list identification.

'@jump:ParsedFormula
'@sub-title Retrieve the CHOICE_FORMULA expression rewritten as nested IF statements
'@return String nested IF expression.
Public Property Get parsedFormula() As String: End Property

'@jump:Categories
'@sub-title Enumerate the textual labels referenced by the CHOICE_FORMULA
'@details Each element represents a label shown to users for a given branch.
'@return BetterArray containing quoted labels.
Public Property Get Categories() As BetterArray: End Property

'@jump:Valid
'@sub-title Determine whether the stored expression is a valid CHOICE_FORMULA
'@details Implementations should verify the token, parenthesis, and argument pairs.
'@return Boolean True when the specification can be parsed.
Public Property Get Valid() As Boolean: End Property

'@jump:ChoiceName
'@sub-title Retrieve the name of the referenced choice list
'@return String identifier for the choice list.
Public Property Get choiceName() As String: End Property
