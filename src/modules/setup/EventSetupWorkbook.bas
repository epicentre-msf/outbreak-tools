Attribute VB_Name = "EventSetupWorkbook"
Option Explicit

Private mBooting As Boolean

'@Folder("Setup")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared EventsSetup service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Sub Workbook_Open()
    Application.EnableEvents = False
    mBooting = True

    On Error GoTo Clean

    SetupEventsManager.WorkbookOpened

Clean:
    mBooting = False
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)

    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If sh.Name = "__checkRep" Then Exit Sub

    Application.EnableEvents = False
    mBooting = True

    On Error GoTo Clean

    SetupEventsManager.SheetActivated sh

Clean:
    mBooting = False
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)

    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If sh.Name = "__checkRep" Then Exit Sub

    Application.EnableEvents = False
    mBooting = True

    On Error GoTo Clean

    SetupEventsManager.SheetChanged sh, Target

Clean:
    mBooting = False
    Application.EnableEvents = True
End Sub
