VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for cached formula metadata"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Contract for cached formula metadata")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IFormulaData
'@description
'Defines the public contract for the FormulaData class, exposing
'membership lookups for allowed Excel functions, permitted separator
'characters, and grouped-formula metadata used by the Formulas parser.


'@section Lookups
'===============================================================================
'@description Membership tests against cached function, character, and grouped-formula registries.

'@jump:special-character-includes
'@sub-title Test whether a character is part of the allowed separators list
'@details
'Implementations check the in-memory character cache built from the
'T_ascii table. Returns False for empty input without accessing
'the cache.
'@param specialChar String. Character to inspect.
'@return Boolean. True when the character exists in the lookup.
Public Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
End Function

'@jump:excel-formulas-includes
'@sub-title Test whether an Excel function is whitelisted
'@details
'Implementations normalise the input to uppercase and check the
'in-memory function cache built from the T_XlsFonctions table.
'Returns False for empty input.
'@param ExcelFormula String. Function name to inspect.
'@return Boolean. True when the function exists in the lookup.
Public Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
End Function

'@jump:group-function-includes
'@sub-title Determine whether a setup function is registered as grouped
'@details
'Implementations scan the grouped-formula registry for a matching
'token. Grouped functions receive special handling during formula
'evaluation and produce aggregated Excel expressions.
'@param functionName String. Grouped token to inspect.
'@return Boolean. True when the function is recognised for grouped parsing.
Public Function IsGroupFunction(ByVal functionName As String) As Boolean
End Function

'@jump:group-aggregator
'@sub-title Retrieve the Excel aggregator mapped to the grouped token
'@details
'Returns the canonical Excel aggregation function (e.g. "SUMIFS",
'"AVERAGE") linked to the provided grouped token. Returns
'vbNullString when the token is not registered.
'@param functionName String. Grouped token to translate.
'@return String. Aggregator name or vbNullString when the token is unknown.
Public Function GroupAggregator(ByVal functionName As String) As String
End Function

'@jump:group-uses-native
'@sub-title Indicate whether the grouped token should emit a native Excel *IFS function
'@details
'When True, the Formulas class emits the native *IFS aggregator
'(e.g. SUMIFS) directly rather than wrapping it in an IF block.
'@param functionName String. Grouped token to inspect.
'@return Boolean. True when the grouped function maps directly to a native Excel *IFS function.
Public Function GroupUsesNativeFunction(ByVal functionName As String) As Boolean
End Function