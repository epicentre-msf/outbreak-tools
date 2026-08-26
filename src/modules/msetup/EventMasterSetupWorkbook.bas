Attribute VB_Name = "EventMasterSetupWorkbook"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Thin workbook-level event handlers delegating to MasterSetupEventsManager.")
'@depends MasterSetupEventsManager
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'This is the code the master setup workbook carries in its ThisWorkbook
'module, the way EventSetupWorkbook is carried by the setup workbook. Every
'handler stays thin and forwards to MasterSetupEventsManager, which is the
'one place master setup events are handled.

Private reentrant As Boolean

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    reentrant = True

    On Error GoTo Clean
    MasterSetupEventsManager.MsWorkbookOpened

Clean:
    'A failed open is otherwise invisible, and the next change event meets
    'the same fault with nothing said.
    If Err.Number <> 0 Then Debug.Print "Workbook_Open error "; Err.Number; " "; Err.Description
    reentrant = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    MasterSetupEventsManager.DisposeMasterSetup
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If reentrant Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    reentrant = True

    On Error GoTo Clean
    MasterSetupEventsManager.MsSheetChanged sh, target

Clean:
    reentrant = False
End Sub
