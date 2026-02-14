Attribute VB_Name = "EventSetupWorkbook"
Option Explicit

Private mBooting As Boolean

'@Folder("Setup")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared EventSetup service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    mBooting = True

    On Error GoTo Clean
    SetupEventsManager.WorkbookOpened

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    Application.ScreenUpdating = False


    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If sh.Name = "__checkRep" Then Exit Sub

    mBooting = True

    On Error GoTo Clean

    SetupEventsManager.SheetActivated sh

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Application.ScreenUpdating = False


    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If sh.Name = "__checkRep" Then Exit Sub

    mBooting = True

    On Error GoTo Clean

    SetupEventsManager.SheetChanged sh, Target

Clean:
    mBooting = False
End Sub
