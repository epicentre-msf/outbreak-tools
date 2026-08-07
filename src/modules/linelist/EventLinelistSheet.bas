Attribute VB_Name = "EventLinelistSheet"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Worksheet-level event template injected into HList sheets during build")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'This module serves as a source template during linelist generation.
'Its code is read by CodeTransfer.TransferWorksheetCode and written into
'the code module of each HList worksheet in the output workbook.
'The module itself is NOT transferred to the output workbook.
'
'It carries the whole worksheet event surface of a data entry sheet. A worksheet
'code module holds as many handlers as it likes, so both handlers travel in one
'TransferWorksheetCode call.

Private mBooting As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If mBooting Then Exit Sub
    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SelectionChanged Me, Target

Clean:
    mBooting = False
End Sub

'The list-auto update runs when the user leaves a data entry sheet. Only an edit
'on such a sheet raises the flag it reads, so the workbook-level handler that
'used to do this made every other sheet of the workbook pay for the check.
Private Sub Worksheet_Deactivate()
    If mBooting Then Exit Sub
    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SheetDeactivated Me

Clean:
    mBooting = False
End Sub
