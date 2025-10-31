Attribute VB_Name = "EventSetupWorkbook"
Option Explicit

'@Folder("Setup")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared EventsSetup service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Sub Workbook_Open()
    SetupEventsManager.WorkbookOpened
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    SetupEventsManager.SheetActivated sh
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    SetupEventsManager.SheetChanged sh, Target
End Sub
