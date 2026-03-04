VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for core analysis worksheet management"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for core analysis worksheet management")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IAnalysis
'@description
'Defines the public contract for the Analysis class: worksheet access, row
'management (add, insert, delete, remove), data exchange (import, export,
'translate, sort, clean), table validation, and checking/trace access.


'@section Elements
'===============================================================================

'@jump:wksh
'@prop-title Worksheet hosting the analysis tables
'@details
'Returns the worksheet bound to this analysis controller during construction.
'@return Worksheet. The backing analysis worksheet.
Public Property Get Wksh() As Worksheet: End Property


'@section RowManagement
'===============================================================================
'@description Add, insert, delete, and remove rows in analysis ListObjects.

'@jump:addrows
'@sub-title Add rows to analysis tables based on the selection header
'@details
'Reads the header instruction in cell A1 of the analysis worksheet and appends
'rows to the matching analysis table(s). Delegates to ManageRows internally.
Public Sub AddRows(): End Sub

'@jump:insertrows
'@sub-title Insert rows based on a worksheet selection
'@details
'Determines which analysis table hosts the selection and mirrors the selection
'height inside that table. Worksheet row insertion is enabled by default to
'protect stacked ListObjects.
'@param targetCell Range. Anchor describing the rows to mirror.
'@param insertShift Optional Boolean. When True, worksheet rows are inserted
'to keep stacked tables aligned. Defaults to True.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = True)
End Sub

'@jump:deleterows
'@sub-title Delete rows intersecting the selection
'@details
'Locates the analysis table hosting the selection and removes the intersecting
'rows, optionally including identifier columns and forcing row shifts.
'@param targetCell Range. Anchor describing the rows to delete.
'@param includeIds Optional Boolean. When True, identifiers are cleared along
'with data. Defaults to True.
'@param forceShift Optional Boolean. When True, worksheet rows are shifted up
'after deletion. Defaults to True.
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = True)
End Sub

'@jump:removerows
'@sub-title Remove rows from analysis tables while preserving required rows
'@details
'Calls the row manager in deletion mode to trim each table down to its required
'minimum row count.
Public Sub RemoveRows(): End Sub

'@jump:managerows
'@sub-title Central row-management dispatcher
'@details
'Reads the header instruction and adjusts analysis tables accordingly, either
'by adding or deleting rows depending on the del flag.
'@param del Optional Boolean. When True, rows are removed; otherwise rows are
'added. Defaults to False.
Public Sub ManageRows(Optional ByVal del As Boolean = False): End Sub


'@section DataExchange
'===============================================================================
'@description Import, export, translate, sort, clean, and validate operations.

'@jump:import
'@sub-title Import analysis ListObjects from another worksheet
'@details
'Copies each ListObject and its associated headers from the source worksheet
'into the managed worksheet, preserving cell formatting and validating the
'result afterward.
'@param sourcesheet Worksheet. The worksheet providing the source analysis tables.
Public Sub Import(ByVal sourcesheet As Worksheet): End Sub

'@jump:export
'@sub-title Export the analysis worksheet to a workbook
'@details
'Duplicates the managed worksheet into the provided workbook so external
'processes can consume the analysis tables.
'@param sourcewb Workbook. The workbook receiving the exported worksheet.
'@param Hide Optional Long. Excel visibility constant. Defaults to xlSheetHidden.
Public Sub Export(ByVal sourcewb As Workbook, Optional ByVal Hide As Long = xlSheetHidden): End Sub

'@jump:translate
'@sub-title Translate analysis labels using a translation object
'@details
'Pushes known textual columns through the translation object so the analysis
'worksheet reflects the currently selected language.
'@param translator ITranslationObject. The translation engine to use.
Public Sub Translate(ByVal translator As ITranslationObject): End Sub

'@jump:sort
'@sub-title Sort analysis ListObjects and enforce minimum rows
'@details
'Iterates through the known analysis tables, sorts them with the correct
'priorities and trims trailing rows so every table meets its structural rules.
Public Sub Sort(): End Sub

'@jump:clean
'@sub-title Clean the analysis worksheet
'@details
'Iterates through all ListObjects on the worksheet and cleans each one via
'CustomTable.Clean.
Public Sub Clean(): End Sub

'@jump:checktables
'@sub-title Ensure all expected analysis tables are present
'@details
'Validates that every expected ListObject exists on the analysis worksheet.
'Raises an error or logs a warning for any that are missing.
Public Sub EnsureAnalysisTables(): End Sub


'@section Checkings
'===============================================================================
'@description Trace and diagnostic access.

'@jump:hascheckings
'@prop-title Whether trace entries were recorded
'@details
'Provides a quick way for callers to know if diagnostics are available.
'@return Boolean. True when at least one trace entry has been logged.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checkingvalues
'@prop-title Retrieve the collected trace entries
'@details
'Returns the underlying checking object when traces are available.
'@return Object. The IChecking instance, or Nothing when no traces exist.
Public Property Get CheckingValues() As Object: End Property
