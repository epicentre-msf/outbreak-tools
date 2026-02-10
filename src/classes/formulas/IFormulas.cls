VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for parsed setup formulas"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Contract for parsed setup formulas")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IFormulas
'@description
'Defines the public contract for the Formulas class, exposing
'read-only access to parsing results, validity checks, and derived
'Excel expressions for both linelist and analysis contexts.
'@depends BetterArray, Checking, IFormulaCondition


'@section Elements
'===============================================================================
'@description Parsed formula metadata and variable lists.

'@jump:varlists
'@prop-title Retrieve the variables referenced by the setup formula
'@details
'Returns a clone of the variable names detected during tokenisation.
'The formulaType parameter controls whether the formula is evaluated
'in an analysis or simple context.
'@param formulaType Optional String. Context (analysis/simple).
'@return BetterArray. Copy of variable names.
Public Property Get Varlists(Optional ByVal formulaType As String = "analysis") As BetterArray
End Property

'@jump:formula-table-list
'@prop-title Retrieve the tokenised representation of the formula
'@details
'Returns a clone of the ordered token list produced during parsing.
'@return BetterArray. Ordered tokens.
Public Property Get FormulaTableList() As BetterArray
End Property

'@jump:has-setup-variables
'@prop-title Determine whether the formula contains literal values
'@details
'Returns True when numeric or text literals were encountered during
'tokenisation.
'@return Boolean. True when literals are present.
Public Property Get HasSetupVariables() As Boolean
End Property

'@jump:is-grouped
'@prop-title Indicate whether the parsed formula targets grouped evaluation
'@details
'Returns "Yes" when the formula uses a grouped function (e.g.
'GROUP_SUM), "No" otherwise.
'@return String. "Yes" for grouped formulas, "No" otherwise.
Public Property Get IsGrouped() As String
End Property

'@jump:has-checkings
'@prop-title Determine whether parsing recorded diagnostic messages
'@details
'Returns True when at least one checking entry was logged during
'evaluation.
'@return Boolean. True when internal checkings exist.
Public Property Get HasChecking() As Boolean
End Property

'@jump:checking-values
'@prop-title Retrieve the internal checking log
'@details
'Returns the IChecking store containing accumulated diagnostics
'logged during evaluation.
'@return IChecking. Accumulated diagnostics.
Public Property Get CheckingValues() As IChecking
End Property

'@jump:reason
'@prop-title Provide the last computed validity explanation
'@details
'Returns a human-readable message describing the validity outcome
'for the requested context. Returns the default success message when
'the formula is valid.
'@param formulaType Optional String. Context (analysis/simple).
'@return String. Message describing the validity outcome.
Public Property Get Reason(Optional ByVal formulaType As String = "analysis") As String
End Property


'@section Operations
'===============================================================================
'@description Validation and Excel formula generation.

'@jump:valid
'@prop-title Determine whether the formula is valid for the supplied context
'@details
'Triggers a lazy evaluation for the requested context and returns
'the cached validity flag.
'@param formulaType Optional String. Context (analysis/simple).
'@return Boolean. Validity flag.
Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
End Property

'@jump:parsed-linelist-formula
'@prop-title Build the Excel expression for a linelist worksheet
'@details
'Replaces variable tokens with worksheet addresses or structured
'references depending on the useTableName flag.
'@param useTableName Optional Boolean. True to use structured references.
'@param tablePrefix Optional String. Prefix for structured references.
'@return String. Excel formula text.
Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
End Property

'@jump:parsed-analysis-formula
'@prop-title Build the Excel expression for an analysis worksheet
'@details
'Replaces variable tokens with conditional aggregation fragments
'supplied by the IFormulaCondition provider.
'@param formCond IFormulaCondition. Provides aggregator behaviour.
'@param tablePrefix Optional String. Prefix for structured references.
'@param Connector Optional String. Connector between conditions.
'@return String. Excel formula text.
Public Property Get ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
End Property
