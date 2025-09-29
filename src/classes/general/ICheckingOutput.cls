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

Option Explicit

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of checking output")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument


'@section Overview
'@details Declares the surface area required to print checking results and apply worksheet filters.


'@jump:create


'@jump:PrintOutput
'@sub-title Write checking values to a worksheet
'@details Implementations must write the supplied check results to the bound worksheet, initialising formatting and filters.
'@param checkTable BetterArray containing IChecking instances
Public Sub PrintOutput(ByVal checkTable As BetterArray): End Sub


'@jump:FilterWorksheet
'@sub-title Apply the active filter on the output worksheet
'@details Re-evaluates the drop-down filter to hide or reveal rows matching the requested severity.
'@param selectionValue Optional String value matching the filter (All/Errors/Warnings/Notes/Infos)
Public Sub FilterWorksheet(Optional ByVal selectionValue As String = vbNullString): End Sub
