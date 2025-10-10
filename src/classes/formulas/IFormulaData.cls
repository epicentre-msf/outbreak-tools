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
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
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

'@label:IsGroupFunction
'@jump:IsGroupFunction
'@sub-title Determine whether a setup function is registered as grouped
'@param functionName String grouped token to inspect.
'@return Boolean True when the function is recognised for grouped parsing.
Public Function IsGroupFunction(ByVal functionName As String) As Boolean
End Function

'@label:GroupAggregator
'@jump:GroupAggregator
'@sub-title Retrieve the Excel aggregator mapped to the grouped token
'@param functionName String grouped token to translate.
'@return String aggregator name or vbNullString when the token is unknown.
Public Function GroupAggregator(ByVal functionName As String) As String
End Function

'@label:GroupUsesNativeFunction
'@jump:GroupUsesNativeFunction
'@sub-title Indicate whether the grouped token should emit a native Excel *IFS function
'@param functionName String grouped token to inspect.
'@return Boolean True when the grouped function maps directly to a native Excel *IFS function.
Public Function GroupUsesNativeFunction(ByVal functionName As String) As Boolean
End Function
