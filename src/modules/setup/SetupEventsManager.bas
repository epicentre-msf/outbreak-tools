Attribute VB_Name = "SetupEventsManager"
Option Explicit

'@Folder("Setup")
'@ModuleDescription("Workbook-level event wrappers delegating to the EventSetup service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private setupService As IEventSetup

'@section Service lifecycle
'===============================================================================
Private Function Service() As IEventSetup
    If setupService Is Nothing Then
        Set setupService = EventSetup.Create(ThisWorkbook)
    End If
    Set Service = setupService
End Function

Public Sub ResetEventSetupCaches()
    If Not setupService Is Nothing Then
        setupService.ResetCaches
    End If
End Sub

Public Sub DisposeEventSetup()
    Set setupService = Nothing
End Sub

'@section Workbook entry points
'===============================================================================
Public Sub WorkbookOpened()
    Dim scope As IApplicationState

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    On Error GoTo Cleanup
        Service.OnWorkbookOpen
Cleanup:
    If Not scope Is Nothing Then scope.Restore
End Sub

Public Sub SheetActivated(ByVal sh As Worksheet)
    Dim scope As IApplicationState

    If sh Is Nothing Then Exit Sub

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    On Error GoTo Cleanup
        Service.OnSheetActivate sh
Cleanup:
    If Not scope Is Nothing Then scope.Restore
End Sub

Public Sub SheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    Dim scope As IApplicationState

    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    On Error GoTo Cleanup
        Service.OnSheetChange sh, target
Cleanup:
    If Not scope Is Nothing Then scope.Restore
End Sub

Public Sub RefreshAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
    Dim scope As IApplicationState

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    On Error GoTo Cleanup
        Service.UpdateAnalysisDropdowns forceUpdate
Cleanup:
    If Not scope Is Nothing Then scope.Restore
End Sub

Public Sub RecalculateAnalysis()
    Dim scope As IApplicationState

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    On Error GoTo Cleanup
        Service.RecalculateAnalysis
Cleanup:
    If Not scope Is Nothing Then scope.Restore
End Sub

Public Sub ResetTranslationCounter()
    Service.ResetTranslationCounter
End Sub

Public Function EventSetupService() As IEventSetup
    Set EventSetupService = Service()
End Function
