Attribute VB_Name = "EventLinelistWorkbook"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared LinelistEventsManager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private mBooting As Boolean

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.WorkbookOpened

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SheetActivated sh

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SheetDeactivated sh

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.SheetChanged sh, target

Clean:
    mBooting = False
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As Range, Cancel As Boolean)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.DoubleClicked sh, target

Clean:
    mBooting = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.FormatStaleValues = True
    LinelistEventsManager.DisposeEventLinelist
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Application.CalculateBeforeSave = False

    'Guard: restrict Save As to macro-enabled binary format (.xlsb) only
    If SaveAsUI Then
        Dim savePath As Variant

        Cancel = True
        savePath = Application.GetSaveAsFilename( _
            InitialFileName:=ThisWorkbook.Name, _
            FileFilter:="Excel Binary Workbook (*.xlsb), *.xlsb", _
            Title:="Save As")

        If savePath <> False Then
            On Error Resume Next
            Application.EnableEvents = False
            ThisWorkbook.SaveAs Filename:=CStr(savePath), FileFormat:=xlExcel12
            Application.EnableEvents = True
            On Error GoTo 0
        End If
    End If
End Sub
