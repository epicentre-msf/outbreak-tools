Attribute VB_Name = "EventMasterSetupWorkbook"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Thin workbook-level event handlers delegating to MasterSetupEventsManager.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'This is the code the master setup workbook carries in its ThisWorkbook
'module, the way EventSetupWorkbook is carried by the setup workbook. Every
'handler stays thin and forwards to MasterSetupEventsManager, which is the
'one place master setup events are handled.

Private mBooting As Boolean

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    mBooting = True

    On Error GoTo Clean
    MasterSetupEventsManager.MsWorkbookOpened

Clean:
    mBooting = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    MasterSetupEventsManager.DisposeMasterSetup
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    MasterSetupEventsManager.MsSheetChanged sh, target

Clean:
    mBooting = False
End Sub
