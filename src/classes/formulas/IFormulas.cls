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

'@Folder("Formulas")
'@ModuleDescription("Contract for parsed setup formulas")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Interface
'@depends BetterArray, Checking, IFormulaCondition

Option Explicit

'@section Overview
'===============================================================================
'@details Exposes read-only access to parsing results, validity, and derived
'@details Excel expressions for both linelist and analysis contexts.

'@section Elements
'===============================================================================

'@jump:Varlists
'@sub-title Retrieve the variables referenced by the setup formula
'@param formulaType Optional String context (analysis/simple).
'@return BetterArray copy of variable names.
Public Property Get Varlists(Optional ByVal formulaType As String = "analysis") As BetterArray
End Property

'@jump:FormulaTableList
'@sub-title Retrieve the tokenised representation of the formula
'@return BetterArray of ordered tokens.
Public Property Get FormulaTableList() As BetterArray
End Property

'@jump:HasSetupVariables
'@sub-title Determine whether the formula contains literal values
'@return Boolean True when literals are present.
Public Property Get HasSetupVariables() As Boolean
End Property

'@jump:IsGrouped
'@sub-title Indicate whether the parsed formula targets grouped evaluation
'@return String "Yes" for grouped formulas, "No" otherwise.
Public Property Get IsGrouped() As String
End Property

'@jump:HasChecking
'@sub-title Determine whether parsing recorded diagnostic messages
'@return Boolean True when internal checkings exist.
Public Property Get HasChecking() As Boolean
End Property

'@jump:CheckingValues
'@sub-title Retrieve the internal checking log
'@return IChecking containing accumulated diagnostics.
Public Property Get CheckingValues() As IChecking
End Property

'@jump:Reason
'@sub-title Provide the last computed validity explanation
'@param formulaType Optional String context (analysis/simple).
'@return String message describing the validity outcome.
Public Property Get Reason(Optional ByVal formulaType As String = "analysis") As String
End Property

'@section Operations
'===============================================================================

'@jump:Valid
'@sub-title Determine whether the formula is valid for the supplied context
'@param formulaType Optional String context (analysis/simple).
'@return Boolean validity flag.
Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
End Property

'@jump:ParsedLinelistFormula
'@sub-title Build the Excel expression for a linelist worksheet
'@param useTableName Optional Boolean True to use structured references.
'@param tablePrefix Optional String prefix for structured references.
'@return String Excel formula text.
Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
End Property

'@jump:ParsedAnalysisFormula
'@sub-title Build the Excel expression for an analysis worksheet
'@param formCond IFormulaCondition providing aggregator behaviour.
'@param tablePrefix Optional String prefix for structured references.
'@param Connector Optional String connector between conditions.
'@return String Excel formula text.
Public Property Get ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
End Property

