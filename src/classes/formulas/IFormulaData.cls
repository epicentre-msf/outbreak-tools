VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of FormulaData"

'@Folder("Formula")
'@ModuleDescription("Contract for cached formula metadata")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Interface

Option Explicit

'@section Overview
'===============================================================================
'@details Exposes membership lookups for both allowed Excel functions and
'@details permitted separator characters used by the Formulas parser.

'@section Lookups
'===============================================================================

'@label:SpecialCharacterIncludes
'@jump:SpecialCharacterIncludes
'@sub-title Test whether a character is part of the allowed separators list
'@param specialChar String character to inspect.
'@return Boolean True when the character exists in the lookup.
Public Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
End Function

'@label:ExcelFormulasIncludes
'@jump:ExcelFormulasIncludes
'@sub-title Test whether an Excel function is whitelisted
'@param ExcelFormula String function name to inspect.
'@return Boolean True when the function exists in the lookup.
Public Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
End Function

