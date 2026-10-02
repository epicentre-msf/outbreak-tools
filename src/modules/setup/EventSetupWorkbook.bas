Attribute VB_Name = "EventSetupWorkbook"
Option Explicit

Private reentrant As Boolean

'@Folder("Setup")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared EventSetup service")
'@depends EventsManager
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'The four sheets SetupPreparation.WatchedSheetNames registers a watcher on. A
'committed edit on any other sheet in this workbook has nothing to record and
'nothing to recalculate, so it is dropped here before any Application property
'is written. The checking report sheet is one of those: SetupErrors writes to it
'while events are live, and it must never come back through this handler.
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_CHECKING As String = "__checkRep"

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    reentrant = True

    On Error GoTo Clean
    EventsManager.WorkbookOpened

Clean:
    reentrant = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'FormatStaleValues is an Application setting, so leaving it off would follow
    'the user into every other workbook of the session.
    On Error Resume Next
    Application.FormatStaleValues = True
    EventsManager.DisposeEventSetup
    On Error GoTo 0
End Sub

'Screen updating is left alone here. Excel turns it back on by itself when the
'handler returns, and the False that used to open this routine therefore made
'every sheet the user moved to repaint the window twice. EventsManager takes
'the busy state instead, and only for the one sheet whose activation does work
'worth hiding.
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    If reentrant Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If sh.Name = SHEET_CHECKING Then Exit Sub

    reentrant = True

    On Error GoTo Clean

    EventsManager.SheetActivated sh

Clean:
    reentrant = False
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    If reentrant Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    Select Case sh.Name
    Case SHEET_DICTIONARY, SHEET_CHOICES, SHEET_EXPORTS, SHEET_ANALYSIS
    Case Else
        Exit Sub
    End Select

    reentrant = True

    On Error GoTo Clean

    EventsManager.SheetChanged sh, Target

Clean:
    reentrant = False
End Sub
