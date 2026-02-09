VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICheckingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of checking output"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of checking output")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument

Option Explicit

'@class ICheckingOutput
'@description
'Defines the public contract for the CheckingOutput class, exposing
'worksheet binding, diagnostic entry rendering, and interactive
'filtering for severity and title criteria.


'@section Elements
'===============================================================================
'@description Worksheet binding and output metadata.

'@jump:wksh
'@prop-title Retrieve the worksheet backing this output writer
'@details
'Returns the worksheet that receives the checking output.
'@return Worksheet. Target worksheet reference.
Public Property Get Wksh() As Worksheet: End Property

'@jump:name
'@prop-title Name of the output session
'@details
'Returns the label assigned during creation.
'@return String. Output session label.
Public Property Get Name() As String: End Property


'@section Operations
'===============================================================================
'@description Rendering and filtering of diagnostic entries.

'@jump:print-output
'@sub-title Write checking values to a worksheet
'@details
'Implementations must write the supplied check results to the bound
'worksheet, initialising formatting and filters as needed.
'@param checkTable BetterArray. Collection of IChecking instances to print.
Public Sub PrintOutput(ByVal checkTable As BetterArray): End Sub

'@jump:filter-worksheet
'@sub-title Apply the active filter on the output worksheet
'@details
'Re-evaluates the drop-down filter to hide or reveal rows matching
'the requested severity and/or title.
'@param selectionValue Optional String. Severity filter token (All/Errors/Warnings/Notes/Infos).
'@param titleSelection Optional String. Title limiting rows to a specific checking topic.
Public Sub FilterWorksheet(Optional ByVal selectionValue As String = vbNullString, _
                           Optional ByVal titleSelection As String = vbNullString): End Sub

'@jump:ensure-worksheet-change-handler
'@sub-title Inject the worksheet change handler for interactive filtering
'@details
'Installs VBA code in the worksheet module so filter drop-down
'changes trigger row visibility updates automatically.
Public Sub EnsureWorksheetChangeHandler(): End Sub
