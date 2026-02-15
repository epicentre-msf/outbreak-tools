Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, IDesignerPreparation, DesignerEntry, IDesignerEntry, RibbonDev, LLGeo, ILLGeo, ApplicationState, IApplicationState
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Non-core ribbon logics are callbacks whose absence will not fire a
'warning at workbook opening on the designer. They only execute in
'response to explicit user actions (onAction), never at ribbon load
'time (getLabel, getPressed, getVisible).

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_MAIN As String = "Main"
Private Const PROMPT_TITLE As String = "Designer"


'@section Dev group callbacks
'===============================================================================

'@Description("Initialise the designer workbook: import translations, hide sheets, seed flags.")
'@EntryPoint
Public Sub clickDevInitialize(ByRef control As IRibbonControl)
    Dim prep As IDesignerPreparation
    Dim appScope As IApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set prep = DesignerPreparation.Create(ThisWorkbook)
    prep.Prepare RibbonDev.EnsureDevelopment()

    appScope.Restore
    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    On Error Resume Next
    Application.Cursor = xlDefault
    On Error GoTo 0
    If Err.Number <> 0 Then
        Debug.Print "clickDevInitialize: "; Err.Number; Err.Description
        MsgBox "Unable to initialise designer: " & Err.Description, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub


'@section Manage group callbacks
'===============================================================================

'@Description("Clear all geobase data from the Geo worksheet.")
'@EntryPoint
Public Sub clickDelGeo(ByRef control As IRibbonControl)
    Dim geoSheet As Worksheet
    Dim geo As ILLGeo
    Dim appScope As IApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set geoSheet = ThisWorkbook.Worksheets(SHEET_GEO)
    Set geo = LLGeo.Create(geoSheet)
    geo.Clear

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    On Error Resume Next
    Application.Cursor = xlDefault
    On Error GoTo 0
    If Err.Number <> 0 Then
        Debug.Print "clickDelGeo: "; Err.Number; Err.Description
        MsgBox "Unable to clear geobase: " & Err.Description, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub

'@Description("Clear all entry input ranges on the Main sheet.")
'@EntryPoint
Public Sub clickClearEnt(ByRef control As IRibbonControl)
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))
    entry.Clear

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    On Error Resume Next
    Application.Cursor = xlDefault
    On Error GoTo 0
    If Err.Number <> 0 Then
        Debug.Print "clickClearEnt: "; Err.Number; Err.Description
        MsgBox "Unable to clear entries: " & Err.Description, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub
