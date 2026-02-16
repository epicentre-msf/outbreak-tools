VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for conditions used to create formulas for analysis"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Contract for conditions used to create formulas for analysis")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IFormulaCondition
'@description
'Defines the public contract for the FormulaCondition class, exposing
'variable/condition accessors, dictionary-based validation, and methods
'for rendering Excel predicates used in analysis formulas.


'@section DataAccess
'===============================================================================
'@description Read-only access to the stored variable and condition collections.

'@jump:variables
'@prop-title Retrieve a clone of the stored variable names
'@details
'Returns a BetterArray copy of the variable identifiers in the order
'they were supplied at creation.
'@return BetterArray. The variable names preserving the original order.
Public Property Get Variables() As BetterArray: End Property

'@jump:conditions
'@prop-title Retrieve a clone of the stored condition fragments
'@details
'Returns a BetterArray copy of the predicate fragments in the order
'they were supplied at creation.
'@return BetterArray. The condition fragments preserving the original order.
Public Property Get Conditions() As BetterArray: End Property

'@jump:variables-table
'@prop-title Resolve the table name shared by all tracked variables
'@details
'Returns the memoised table name from a previous validation, or
'triggers a fresh validation pass against the supplied dictionary
'when no cached value exists.
'@param dict ILLdictionary. The dictionary used when validation has not occurred yet.
'@return String. Table identifier or empty string when unavailable.
Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String: End Property


'@section Validation
'===============================================================================
'@description Validation and predicate rendering operations.

'@jump:valid
'@sub-title Validate variable/condition lists against the dictionary
'@details
'Implementations verify that all variables exist in the dictionary and
'belong to the same table. When a tablename is supplied it overrides
'the table resolved from the first variable.
'@param dict ILLdictionary. The metadata provider used for variable lookups.
'@param tablename Optional String. When supplied, forces the expected table name. Defaults to vbNullString.
'@return Boolean. True when validation succeeds.
Public Function Valid(ByVal dict As ILLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean: End Function

'@jump:condition-string
'@sub-title Wrap the predicate in an IF expression referencing the supplied variable
'@details
'Joins all variable/condition pairs using ConditionPredicate and wraps
'the result in an IF function that returns the variable value when the
'predicate evaluates to True.
'@param tablename String. Structured reference table name.
'@param varName String. Variable/column name to return when the condition is met.
'@param Connector Optional String. Connector used between predicates. Defaults to "*".
'@return String. An IF expression or empty string when the predicate is empty.
Public Function ConditionString(ByVal tablename As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String
End Function

'@jump:condition-predicate
'@sub-title Join the variable predicates using the supplied connector
'@details
'Returns the joined predicate without the IF wrapper, useful when
'callers need the raw predicate for embedding in other constructs.
'@param tablename String. Structured reference table name.
'@param Connector Optional String. Connector between predicate fragments. Defaults to "*".
'@return String. Concatenated predicate or empty string when inputs are missing.
Public Function ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
End Function


'@section Diagnostics
'===============================================================================
'@description Exposure of validation traces for diagnostics.

'@jump:has-internal-checks
'@sub-title Determine whether validation recorded diagnostic messages
'@details
'Returns True when at least one diagnostic entry was captured during
'the last validation run.
'@return Boolean. True when internal checks exist.
Public Function HasCheckings() As Boolean
End Function

'@jump:checking-values
'@prop-title Retrieve the internal checking log
'@details
'Returns the accumulated IChecking instance containing all diagnostic
'messages from the last validation run, or Nothing when no diagnostics
'were recorded.
'@return IChecking. The diagnostic log.
Public Property Get CheckingValues() As IChecking
End Property
