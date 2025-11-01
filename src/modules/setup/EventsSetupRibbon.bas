Attribute VB_Name = "EventsSetupRibbon"

Option Explicit

'@Folder("Events")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"


'@section Table Management: callbacks for group CustomGroupManage
'===============================================================================
'@Description("Resize the listObjects in the current sheet")
'@EntryPoint
Public Sub clickResize(ByRef control As IRibbonControl)
    SetupHelpers.ManageRows sheetName:=ActiveSheet.Name, del:=True
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
    MsgBox "Done!", vbInformation
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
    Dim upVal As IUpdatedValues
   
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

    Set upVal = SetupHelpers.ResolveUpdatedValues()
    upVal.SwitchTagsToNo

Cleanup:
    If sheetUnlocked Then SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    If Not app Is Nothing Then app.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    MsgBox "An error occurred while updating translations.", vbCritical
    Resume Cleanup
End Sub
'@Description("Callback for btnTransChange onAction: translate the setup to a selected language")
'@EntryPoint
Public Sub clickTransSetup(ByRef control As IRibbonControl)
    Dim translationsTable As ListObject
    Dim manager As ISetupTranslationsTable
    Dim languages As BetterArray
    Dim selectedLanguage As String
    Dim translator As ITranslationObject
    Dim app As IApplicationState
    Dim translationsUnlocked As Boolean
    Dim success As Boolean

    Set translationsTable = SetupHelpers.ResolveTranslationsTable
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation
        Exit Sub
    End If

    Set manager = SetupTranslationsTable.Create(translationsTable)
    Set languages = manager.Languages
    If languages Is Nothing Or languages.Length = 0 Then
        MsgBox "No translation languages were found. Add a language column first.", vbExclamation
        Exit Sub
    End If

    selectedLanguage = PromptTranslationLanguage(languages)
    If LenB(selectedLanguage) = 0 Then Exit Sub

    On Error GoTo Handler

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    SetupHelpers.UnProtectSetupSheet TRADSHEETNAME
    translationsUnlocked = True

    Set translator = TranslationObject.Create(translationsTable, selectedLanguage)
    SetupHelpers.ApplySetupTranslation translator

    manager.SwitchDefaultLanguage selectedLanguage

    success = True

Cleanup:
    If translationsUnlocked Then SetupHelpers.ProtectSetupSheet TRADSHEETNAME
    If Not app Is Nothing Then app.Restore
    If success Then MsgBox "Done!", vbInformation
    Exit Sub

Handler:
    Debug.Print "clickTransSetup: "; Err.Number; Err.Description
    success = False
    MsgBox "Failed to translate the setup: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Function PromptTranslationLanguage(ByVal languages As BetterArray) As String
    Dim prompt As String
    Dim idx As Long
    Dim response As Variant
    Dim numericResponse As Double
    Dim selection As Long

    If languages Is Nothing Then Exit Function
    If languages.Length = 0 Then Exit Function

    prompt = "Select the language to translate the setup to:" & vbLf
    For idx = languages.LowerBound To languages.UpperBound
        prompt = prompt & CStr(idx - languages.LowerBound + 1) & ". " & CStr(languages.Item(idx)) & vbLf
    Next idx

    response = Application.InputBox(prompt, "Translate the setup", Type:=1)
    If VarType(response) = vbBoolean Then Exit Function

    numericResponse = CDbl(response)
    If numericResponse <> Int(numericResponse) Then GoTo InvalidSelection
    If numericResponse < 1 Or numericResponse > languages.Length Then GoTo InvalidSelection

    selection = CLng(numericResponse)
    PromptTranslationLanguage = Trim$(CStr(languages.Item(languages.LowerBound + selection - 1)))
    MsgBox "Done!", vbInformation
    Exit Function

InvalidSelection:
    MsgBox "Invalid selection.", vbExclamation
End Function


'@section Import and Export management
'===============================================================================

'@Description("Callback for btnExport onAction: export the current setup to a workbook")
'@EntryPoint
Public Sub clickExport(ByRef control As IRibbonControl)
    Dim service As ISetupImportService
    Dim exportPath As String

    On Error GoTo Handler

    Set service = SetupImportService.Create(ThisWorkbook.FullName)
    service.Export

    exportPath = service.LastExportFile
    If LenB(exportPath) > 0 Then
        MsgBox "Setup exported to:" & vbCrLf & exportPath, vbInformation
    End If
    Exit Sub

Handler:
    Debug.Print "clickExport: "; Err.Number; Err.Description
    MsgBox "Failed to export the setup: " & Err.Description, vbCritical
End Sub

'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickImport(ByRef control As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=False
    [Imports].Show
End Sub


'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickClearSetup(ByRef control As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=True
    [Imports].Show
End Sub

'@Description("Callback for btnImpExp onAction: import setup elements from a workbook using table mode")
'@EntryPoint
Public Sub clickImportFile(ByRef control As IRibbonControl)
    Const SUCCESS_MESSAGE As String = "Workbook import completed."

    Dim importPath As String
    Dim service As ISetupImportService
    Dim pass As IPasswords
    Dim app As IApplicationState
    Dim sheets As BetterArray
    Dim success As Boolean

    On Error GoTo Handler

    importPath = SetupHelpers.SelectSetupImportPath("*.xlsx")
    If LenB(importPath) = 0 Then Exit Sub

    Set service = SetupImportService.Create(importPath)
    Set pass = SetupHelpers.ResolveSetupPasswords()
    Set sheets = SetupHelpers.DefaultSetupSheets()
    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False
    service.Check True, True, True, True, True

    service.ImportFromWorkbook pass, sheets
    success = True

Cleanup:
    If Not app Is Nothing Then app.Restore
    If success Then
        SetupHelpers.PostImportMaintenance SUCCESS_MESSAGE
    End If
    Exit Sub

Handler:
    Debug.Print "clickImportFile: "; Err.Number; Err.Description
    success = False
    MsgBox "Failed to import workbook data: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub clickCheck(ByRef control As IRibbonControl)
    SetupHelpers.CheckTheSetup
End Sub

'@section Formatter
'===============================================================================
Public  Sub clickEditStyle(ByRef control As IRibbonControl)
    Const FORMATSHEET As String = "__formatter"
    Static opened As Boolean
    Dim pass As IPasswords
    Dim targetsheet As Worksheet
    
    On Error GoTo Handler

    Set pass = SetupHelpers.ResolveSetupPasswords()

    pass.UnProtect ThisWorkbook
    Set targetsheet = ThisWorkbook.Worksheets(FORMATSHEET)

    If (Not opened) Then
        ThisWorkbook.Worksheets(FORMATSHEET).Visible = xlSheetVisible
        targetSheet.Activate
    Else
        targetSheet.Visible = xlSheetVeryHidden
    End If
    
    opened = (Not opened)
    pass.Protect ThisWorkbook

Handler:
End Sub

'@section Visibility of some buttons
'===============================================================================
Public Sub SetupButtonVisible(control As IRibbonControl, ByRef returnedVal)
    If (control.Id = "btnDelLoRow") Or (control.Id="btnSort") Then
        returnedVal = CBool((ActiveSheet.Name <> TRADSHEETNAME))
    ElseIf (control.Id = "btnDelLoCol") Then
        returnedVal = CBool((ActiveSheet.Name = TRADSHEETNAME))
    Else
        returnedVal = True
    End If
End Sub


'@section Initializations 
'===============================================================================
'@EntryPoint
'@Description("Initialise development environment - logic provided by consuming workbook")
Public Sub clickDevInitialize(ByRef control As IRibbonControl)
   Dim prep As ISetupPreparation
   Set prep = SetupPreparation.Create(ThisWorkbook)
   prep.Prepare RibbonDev.EnsureDevelopment()
   'Protecting required worksheet
   ProtectSetupSheet ResolveSetupSheetName("dict")
   ProtectSetupSheet ResolveSetupSheetName("choi")
   ProtectSetupSheet ResolveSetupSheetName("ana")
   ProtectSetupSheet ResolveSetupSheetName("exp")
   ProtectSetupSheet ResolveSetupSheetName("trans")
   MsgBox "Done!", vbInformation
End Sub