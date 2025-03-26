Attribute VB_Name = "EventsSetupWorkbook"
Attribute VB_Description = "Events for changes at the workbook level"
Option Explicit

'@ModuleDescription("Events for changes at the workbook level")
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage
'@Folder("Events")

Private Sub Workbook_Open()
    OpenedWorkbook
End Sub