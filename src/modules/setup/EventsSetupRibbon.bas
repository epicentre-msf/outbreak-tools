Attribute VB_Name = "EventsSetupRibbon"

Option Explicit

'@Folder("Events")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const EXPORTSHEETNAME As String = "Exports"


'@section Table Management: callbacks for group CustomGroupManage
'===============================================================================
'@Description("Resize the listObjects in the current sheet")
'@EntryPoint
Public Sub clickResize(ByRef control As IRibbonControl)
    SetupHelpers.ManageRows sheetName:=ActiveSheet.Name del:=True
End Sub

'@Description("add rows to listObject")
'@EntryPoint
Public Sub clickAddRows(ByRef control As Office.IRibbonControl)
    SetupHelpers.ManageRows sheetName:=ActiveSheet.Name, del:=False
End Sub

'@Description("Clear all the filters in the current sheet")
'@EntryPoint
Public Sub clickFilters(ByRef control As IRibbonControl)

    Dim app As IApplicationState
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet

    On Error GoTo Handler
    
    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.ClearSheetFilters targetSheet.Name

Cleanup:
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickFilters: "; Err.Number; Err.Description
    Resume Cleanup
End Sub


'@Description("Sort setup tables depending on active sheet")
'@EntryPoint
Public Sub clickSortTables(ByRef control As IRibbonControl)

    Dim app As IApplicationState
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.SortSetupTables targetSheet.Name

Cleanup:
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickSortTables: "; Err.Number; Err.Description
    Resume Cleanup
End Sub

'@Description("Insert a list row at the active position")
'@EntryPoint
Public Sub clickInsertRow(ByRef control As IRibbonControl)

    Dim app As IApplicationState
    Dim targetSheet As Worksheet
    Dim targetCell As Range

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub
    Set targetCell = ActiveCell

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.InsertListRowAt targetSheet.Name, targetCell

Cleanup:
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickInsertRow: "; Err.Number; Err.Description
    Resume Cleanup
End Sub

'@Description("Delete the current list row when the active cell belongs to a table")
'@EntryPoint
Public Sub clickDelLoRows(ByRef control As IRibbonControl)

    Dim app As IApplicationState
    Dim targetSheet As Worksheet
    Dim targetCell As Range

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub
    Set targetCell = ActiveCell

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.DeleteListRowAt targetSheet.Name, targetCell

Cleanup:
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickDelLoRows: "; Err.Number; Err.Description
    Resume Cleanup
End Sub

'@Description("Delete the current list column when the active cell belongs to a table")
'@EntryPoint
Public Sub clickDelLoColumn(ByRef control As IRibbonControl)

    Dim app As IApplicationState
    Dim targetSheet As Worksheet
    Dim targetCell As Range

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub
    Set targetCell = ActiveCell

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.DeleteListColumnAt targetSheet.Name, targetCell

Cleanup:
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickDelLoColumn: "; Err.Number; Err.Description
    Resume Cleanup
End Sub


'@section Translation Management: callbacks for group CustomGroupTrans
'===============================================================================

'@Description("Callback for editLang onChange: add translation language columns")
'@EntryPoint
Public Sub clickAddLang(ByRef control As IRibbonControl, ByRef text As String)
    Dim languages As String
    Dim answer As VbMsgBoxResult
    Dim translationsTable As ListObject
    Dim manager As ISetupTranslationsTable
    Dim app As IApplicationState
    Dim sheetUnlocked As Boolean
    Dim success As Boolean

    languages = Trim$(text)
    If LenB(languages) = 0 Then Exit Sub

    answer = MsgBox("Do you really want to add language(s) " & languages & " to translations?", vbYesNo + vbQuestion, "Confirm")
    If answer <> vbYes Then Exit Sub

    Set translationsTable = SetupHelpers.ResolveTranslationsTable
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.UnProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = True

    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.EnsureLanguages languages

    SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = False

    success = True

Cleanup:
    If sheetUnlocked Then SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    If Not app Is Nothing Then app.Restore
    If success Then MsgBox "Done!", vbInformation
    Exit Sub

Handler:
    Debug.Print "clickAddLang: "; Err.Number; Err.Description
    success = False
    Resume Cleanup
End Sub

'@Description("Callback for btnTransAdd onAction: update translations from registry")
'@EntryPoint
Public Sub clickAddTrans(ByRef control As IRibbonControl)
    Dim answer As VbMsgBoxResult
    Dim translationsTable As ListObject
    Dim registrySheet As Worksheet
    Dim manager As ISetupTranslationsTable
    Dim app As IApplicationState
    Dim sheetUnlocked As Boolean
    Dim success As Boolean

    answer = MsgBox("Do you want to update the translation sheet?", vbYesNo + vbQuestion, "Confirm")
    If answer <> vbYes Then Exit Sub

    Set translationsTable = SetupHelpers.ResolveTranslationsTable
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation
        Exit Sub
    End If

    Set registrySheet = SetupHelpers.ResolveRegistrySheet
    If registrySheet Is Nothing Then
        MsgBox "Registry sheet was not found.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.UnProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = True

    On Error Resume Next
        translationsTable.AutoFilter.ShowAllData
    On Error GoTo 0

    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.UpdateFromRegistry registrySheet

    SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = False

    EventsGlobal.SetAllUpdatedTo "no"
    success = True

Cleanup:
    If sheetUnlocked Then SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    If Not app Is Nothing Then app.Restore
    If success Then MsgBox "Done!", vbInformation
    Exit Sub

Handler:
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    success = False
    MsgBox "An error occurred while updating translations.", vbCritical
    Resume Cleanup
End Sub

'@section Visibility of some buttons
'===============================================================================
Public Sub DelVisible(control As IRibbonControl, ByRef returnedVal As Boolean)
    If (control.Id = "btnDelLoRow") Then
        returnedVal = (ActiveSheet.Name <> TRADSHEETNAME)
    ElseIf (control.Id = "btnDelLoCol") Then
        returnedVal = (ActiveSheet.Name = TRADSHEETNAME)    
    End If
End Sub