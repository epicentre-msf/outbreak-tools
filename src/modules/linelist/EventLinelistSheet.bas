Attribute VB_Name = "EventLinelistSelection"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Worksheet-level SelectionChange event template injected into HList sheets during build")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'This module serves as a source template during linelist generation.
'Its code is read by CodeTransfer.TransferWorksheetCode and written into
'the code module of each HList worksheet in the output workbook.
'The module itself is NOT transferred to the output workbook.

Private mBooting As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If mBooting Then Exit Sub
    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SelectionChanged Me, Target

Clean:
    mBooting = False
End Sub
