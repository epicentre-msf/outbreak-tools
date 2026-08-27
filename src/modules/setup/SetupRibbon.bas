Attribute VB_Name = "SetupRibbon"

Option Explicit

'@Folder("Events")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used
'@depends EventsManager, EventSetup, SetupHelpers, SetupPreparation, SetupImport, SetupTranslationsTable, UpdatedValues, RibbonDev, TranslationObject, Passwords, BetterArray

'Every callback here reaches the setup service through
'EventsManager.EventSetupService. Row work, sorting, sheet protection and
'the setup translation all live on EventSetup, so a click crosses the ribbon, the
'manager and the service, and nothing in between.
'
'A callback that asks the user a question asks it BEFORE entering busy state. A
'prompt raised over a frozen screen with a busy cursor reads as a hang.

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"


'@section Table Management: callbacks for group CustomGroupManage
'===============================================================================
'@Description("Resize the listObjects in the current sheet")
'@EntryPoint
Public Sub clickResize(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim sheetName As String

    If ActiveSheet Is Nothing Then Exit Sub
    sheetName = ActiveSheet.Name
    Set svc = EventsManager.EventSetupService

    On Error GoTo Cleanup
    EventsManager.EnterBusyState persist:=False
    svc.ManageRows sheetName, del:=True
Cleanup:
    EventsManager.ExitBusyState
End Sub

'@Description("add rows to listObject")
'@EntryPoint
Public Sub clickAddRows(ByRef ribbonControl As Office.IRibbonControl)
    Dim svc As EventSetup
    Dim sheetName As String

    If ActiveSheet Is Nothing Then Exit Sub
    sheetName = ActiveSheet.Name
    Set svc = EventsManager.EventSetupService

    On Error GoTo Cleanup
    EventsManager.EnterBusyState persist:=False
    svc.ManageRows sheetName, del:=False
Cleanup:
    EventsManager.ExitBusyState
End Sub

'@Description("Clear all the filters in the current sheet")
'@EntryPoint
Public Sub clickFilters(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim sheetName As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    sheetName = targetSheet.Name
    Set svc = EventsManager.EventSetupService

    On Error GoTo Handler

    EventsManager.EnterBusyState calculateOnSave:=False

    svc.UnprotectSetupSheet sheetName

    For Each lo In targetSheet.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            On Error Resume Next
                lo.AutoFilter.ShowAllData
            On Error GoTo 0
        End If
    Next lo

    If targetSheet.AutoFilterMode Then
        targetSheet.AutoFilterMode = False
    End If

    svc.ProtectSetupSheet sheetName

Cleanup:
    EventsManager.ExitBusyState

    Exit Sub

Handler:
    Debug.Print "clickFilters: "; Err.Number; Err.Description
    Resume Cleanup
End Sub


'@Description("Sort setup tables depending on active sheet")
'@EntryPoint
Public Sub clickSortTables(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim sheetName As String

    If ActiveSheet Is Nothing Then Exit Sub
    sheetName = ActiveSheet.Name
    Set svc = EventsManager.EventSetupService

    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    svc.SortTables sheetName

Cleanup:
    EventsManager.ExitBusyState
End Sub

'@Description("Insert a list row at the active position")
'@EntryPoint
Public Sub clickInsertRow(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim sheetName As String
    Dim targetCell As Range

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    sheetName = ActiveSheet.Name
    Set targetCell = Selection
    Set svc = EventsManager.EventSetupService

    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    svc.InsertRows sheetName, targetCell

Cleanup:
    EventsManager.ExitBusyState
End Sub

'@Description("Delete the current list row when the active cell belongs to a table")
'@EntryPoint
Public Sub clickDelLoRows(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim sheetName As String
    Dim targetCell As Range

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    'Asked before the screen freezes.
    If MsgBox("Delete the selected rows?" & vbCrLf & "THIS OPERATION IS IRREVERSIBLE.", _
              vbExclamation + vbYesNo, "Delete Rows") <> vbYes Then Exit Sub

    sheetName = ActiveSheet.Name
    Set targetCell = Selection
    Set svc = EventsManager.EventSetupService

    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    svc.DeleteRows sheetName, targetCell

Cleanup:
    EventsManager.ExitBusyState
End Sub

'@Description("Delete the current list column when the active cell belongs to a table")
'@EntryPoint
Public Sub clickDelLoColumn(ByRef ribbonControl As IRibbonControl)
    Dim sheetName As String
    Dim targetCell As Range

    If ActiveSheet Is Nothing Then Exit Sub
    sheetName = ActiveSheet.Name

    'Only the Translations sheet has columns a user may remove.
    If StrComp(sheetName, TRADSHEETNAME, vbTextCompare) <> 0 Then Exit Sub

    'Asked before the screen freezes.
    If MsgBox("Delete the selected Column?" & vbCrLf & "THIS OPERATION IS IRREVERSIBLE.", _
              vbExclamation + vbYesNo, "Delete Column") <> vbYes Then Exit Sub

    Set targetCell = ActiveCell

    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    SetupHelpers.DeleteListColumnAt sheetName, targetCell

Cleanup:
    EventsManager.ExitBusyState
End Sub


'@section Translation Management: callbacks for group CustomGroupTrans
'===============================================================================

Public Sub clickResetTag(ByRef ribbonControl As IRibbonControl)
   Dim prep As SetupPreparation

   On Error GoTo Handler

   EventsManager.EnterBusyState

   Set prep = SetupPreparation.Create(ThisWorkbook)
   prep.ResetUpdatedRegistry

   EventsManager.ExitBusyState

   'Every range is read again on the spot, with the review on: the labels
   'no range produces are listed and the user says whether they go. The
   'update's own summary is the closing word.
   RunTranslationsUpdate reviewUnseen:=True

   Exit Sub
Handler:
    Debug.Print "clickResetTag: "; Err.Number; Err.Description
    EventsManager.ExitBusyState
End Sub

'@Description("Callback for editLang onChange: add translation language columns")
'@EntryPoint
Public Sub clickAddLang(ByRef ribbonControl As IRibbonControl, ByRef text As String)
    Dim svc As EventSetup
    Dim languages As String
    Dim answer As VbMsgBoxResult
    Dim translationsTable As ListObject
    Dim manager As SetupTranslationsTable
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

    Set svc = EventsManager.EventSetupService

    On Error GoTo Handler

    EventsManager.EnterBusyState calculateOnSave:=False

    svc.UnprotectSetupSheet TRADSHEETNAME
    sheetUnlocked = True

    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.EnsureLanguages languages

    svc.ProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = False

    success = True

Cleanup:
    On Error Resume Next
    If sheetUnlocked Then svc.ProtectSetupSheet TRADSHEETNAME
    On Error GoTo 0
    EventsManager.ExitBusyState
    If success Then MsgBox "Done!", vbInformation
    Exit Sub

Handler:
    Debug.Print "clickAddLang: "; Err.Number; Err.Description
    success = False
    Resume Cleanup
End Sub

'@Description("Callback for btnTransAdd onAction: update translations from registry")
'@EntryPoint
Public Sub clickAddTrans(ByRef ribbonControl As IRibbonControl)
    If MsgBox("Do you want to update the translation sheet?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub
    RunTranslationsUpdate reviewUnseen:=False
End Sub

'@sub-title Update the translations table from the registry, the shared body of the two ribbon buttons.
'@details Update Translations runs it as is. Reset tags runs it with the
'review on: the update then offers the labels no range of the setup
'produces and asks whether they go, right on the click.
'@param reviewUnseen Boolean. True asks about the unseen labels.
'@return Boolean. True when the update ran through.
Private Function RunTranslationsUpdate(ByVal reviewUnseen As Boolean) As Boolean
    Dim svc As EventSetup
    Dim translationsTable As ListObject
    Dim registrySheet As Worksheet
    Dim manager As SetupTranslationsTable
    Dim sheetUnlocked As Boolean
    Dim upVal As UpdatedValues

    Set translationsTable = SetupHelpers.ResolveTranslationsTable
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation
        Exit Function
    End If

    Set registrySheet = SetupHelpers.ResolveRegistrySheet
    If registrySheet Is Nothing Then
        MsgBox "Registry sheet was not found.", vbExclamation
        Exit Function
    End If

    Set svc = EventsManager.EventSetupService

    On Error GoTo Handler

    EventsManager.EnterBusyState calculateOnSave:=False

    svc.UnprotectSetupSheet TRADSHEETNAME
    sheetUnlocked = True

    On Error Resume Next
        translationsTable.AutoFilter.ShowAllData
    On Error GoTo Handler

    Set manager = SetupTranslationsTable.Create(translationsTable)
    If reviewUnseen Then manager.RequestUnseenReview
    manager.UpdateFromRegistry registrySheet

    svc.ProtectSetupSheet TRADSHEETNAME
    sheetUnlocked = False

    Set upVal = SetupHelpers.ResolveUpdatedValues()
    upVal.SwitchTagsToNo
    RunTranslationsUpdate = True

Cleanup:
    On Error Resume Next
    If sheetUnlocked Then svc.ProtectSetupSheet TRADSHEETNAME
    On Error GoTo 0
    EventsManager.ExitBusyState
    Exit Function

Handler:
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    MsgBox "An error occurred while updating translations.", vbCritical
    Resume Cleanup
End Function

'@Description("Callback for btnTransChange onAction: translate the setup to a selected language")
'@EntryPoint
Public Sub clickTransSetup(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim translationsTable As ListObject
    Dim manager As SetupTranslationsTable
    Dim languages As BetterArray
    Dim selectedLanguage As String
    Dim translator As TranslationObject
    Dim translationsUnlocked As Boolean
    Dim success As Boolean
    Dim nbMissing As Long
    Dim dupLabels As String

    Set translationsTable = SetupHelpers.ResolveTranslationsTable
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation
        Exit Sub
    End If

    Set svc = EventsManager.EventSetupService

    'Armed before the first unprotect. A failure between the unprotect and the
    'language prompt used to leave the Translations sheet open with no message.
    On Error GoTo Handler

    svc.UnprotectSetupSheet TRADSHEETNAME
    translationsUnlocked = True

    Set manager = SetupTranslationsTable.Create(translationsTable)
    Set languages = manager.Languages
    If (languages Is Nothing) Or (languages.Length = 0) Then
        MsgBox "No translation languages were found. Add a language column first.", vbExclamation
        GoTo Cleanup
    End If

    selectedLanguage = PromptTranslationLanguage(languages)
    If LenB(selectedLanguage) = 0 Then
        GoTo Cleanup
    End If

    'Provide the number of Mission Labels of one specific language
    nbMissing = manager.MissingLabels(selectedLanguage)

    If (nbMissing > 0) Then
        MsgBox "Aborted translation of the setup: Language " & selectedLanguage & _
               " has " & nbMissing & " missing labels. Please fill them before attempting a translation.", vbExclamation
        GoTo Cleanup
    End If

    If manager.DuplicateLabels(dupLabels, selectedLanguage) Then
        MsgBox "Aborted translation of the setup. " & dupLabels, vbExclamation
        GoTo Cleanup
    End If

    EventsManager.EnterBusyState calculateOnSave:=False

    Set translator = TranslationObject.Create(translationsTable, selectedLanguage)
    svc.ApplySetupTranslation translator

    manager.SwitchDefaultLanguage selectedLanguage

    'A translation renames every header the watcher registry and the analysis
    'dropdowns were built from, and every label the analysis formulas look up.
    'Without this the setup came back translated and reading the old language.
    SetupHelpers.PostImportMaintenance

    success = True

Cleanup:
    On Error Resume Next
    If translationsUnlocked Then svc.ProtectSetupSheet TRADSHEETNAME
    On Error GoTo 0
    EventsManager.ExitBusyState
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
    Exit Function

InvalidSelection:
    MsgBox "Invalid selection.", vbExclamation
End Function


'@section Import and Export management
'===============================================================================

'@Description("Callback for btnExport onAction: export the current setup to a workbook")
'@EntryPoint
Public Sub clickExport(ByRef ribbonControl As IRibbonControl)
    Dim svc As EventSetup
    Dim service As SetupImport
    Dim exportPath As String
    Dim analysisSheet As String

    Set svc = EventsManager.EventSetupService

    On Error GoTo Handler

    analysisSheet = SetupHelpers.ResolveSetupSheetName("ana")

    EventsManager.EnterBusyState

    Set service = SetupImport.Create(ThisWorkbook.FullName)

    'UnProtect the analysis before proceeding
    svc.UnprotectSetupSheet analysisSheet
    service.Export
    svc.ProtectSetupSheet analysisSheet

    exportPath = service.LastExportFile

    EventsManager.ExitBusyState

    If LenB(exportPath) > 0 Then
        MsgBox "Setup exported to: " & vbCrLf & exportPath, vbInformation
    End If

    Exit Sub

Handler:
    'Re-protect BEFORE restoring screen state to avoid visible flash
    On Error Resume Next
    svc.ProtectSetupSheet SetupHelpers.ResolveSetupSheetName("ana")
    On Error GoTo 0
    EventsManager.ExitBusyState
    Debug.Print "clickExport: "; Err.Number; Err.Description
    MsgBox "Failed to export the setup: " & Err.Description, vbCritical
End Sub

'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickImport(ByRef ribbonControl As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=False
    [Imports].Show
End Sub


'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickClearSetup(ByRef ribbonControl As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=True
    [Imports].Show
End Sub

'@Description("Callback for btnImpExp onAction: import setup elements from a workbook using table mode")
'@EntryPoint
Public Sub clickImportFile(ByRef ribbonControl As IRibbonControl)
    Dim importPath As String
    Dim service As SetupImport
    Dim pass As Passwords
    Dim sheets As BetterArray
    Dim success As Boolean
    Dim originalSheet As Worksheet

    'The file picker comes first, so nothing is frozen while it is open.
    importPath = SetupHelpers.SelectSetupImportPath("*.xlsx")
    If LenB(importPath) = 0 Then Exit Sub

    On Error GoTo Handler

    Set service = SetupImport.Create(importPath)
    Set pass = SetupHelpers.ResolveSetupPasswords()
    Set sheets = SetupHelpers.BuildSelectedSheets(True, True, True, True, True)
    Set originalSheet = ActiveSheet

    EventsManager.EnterBusyState calculateOnSave:=False

    service.ImportFromWorkbook pass, sheets
    SetupHelpers.PostImportMaintenance
    success = True

Cleanup:
    EventsManager.ExitBusyState
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.ScreenUpdating = True
    If success Then MsgBox "Import Done!"
    Exit Sub

Handler:
    Debug.Print "clickImportFile: "; Err.Number; Err.Description
    success = False
    MsgBox "Failed to import workbook data: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub clickCheck(ByRef ribbonControl As IRibbonControl)
    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    SetupHelpers.CheckTheSetup

Cleanup:
    EventsManager.ExitBusyState
End Sub

'@section Formatter
'===============================================================================
Public Sub clickEditStyle(ByRef ribbonControl As IRibbonControl)
    Const FORMATSHEET As String = "__formatter"

    Dim pass As Passwords
    Dim targetSheet As Worksheet

    On Error GoTo Handler

    Set pass = SetupHelpers.ResolveSetupPasswords()

    pass.UnProtect ThisWorkbook
    Set targetSheet = ThisWorkbook.Worksheets(FORMATSHEET)

    'The sheet carries the open state. A Static variable used to, and any VBA
    'state reset made it disagree with the sheet, so the next click hid a sheet
    'that was already hidden.
    If targetSheet.Visible = xlSheetVisible Then
        targetSheet.Visible = xlSheetVeryHidden
    Else
        targetSheet.Visible = xlSheetVisible
        targetSheet.Activate
    End If

Cleanup:
    On Error Resume Next
    If Not pass Is Nothing Then pass.Protect ThisWorkbook
    On Error GoTo 0
    Exit Sub

Handler:
    Debug.Print "clickEditStyle: "; Err.Number; Err.Description
    Resume Cleanup
End Sub

'@section Visibility of some buttons
'===============================================================================
Public Sub SetupButtonVisible(ribbonControl As IRibbonControl, ByRef returnedVal)
    If (ribbonControl.Id = "btnDelLoRow") Or (ribbonControl.Id = "btnSort") Then
        returnedVal = CBool((ActiveSheet.Name <> TRADSHEETNAME))
    ElseIf (ribbonControl.Id = "btnDelLoCol") Then
        returnedVal = CBool((ActiveSheet.Name = TRADSHEETNAME))
    Else
        returnedVal = True
    End If
End Sub


'@section Initializations
'===============================================================================
'@EntryPoint
'@Description("Initialise development environment - logic provided by consuming workbook")
Public Sub clickDevInitialize(ByRef ribbonControl As IRibbonControl)
   Dim prep As SetupPreparation

   On Error GoTo Cleanup

   EventsManager.EnterBusyState

   Set prep = SetupPreparation.Create(ThisWorkbook)
   prep.Prepare RibbonDev.EnsureDevelopment()

   EventsManager.ExitBusyState
   MsgBox "Done!", vbInformation
   Exit Sub

Cleanup:
   EventsManager.ExitBusyState
   Debug.Print "clickDevInitialize: "; Err.Number; Err.Description
End Sub
