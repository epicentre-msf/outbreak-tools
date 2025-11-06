Attribute VB_Name = "EventsMasterSetupRibbon"
Option Explicit

'@Folder("Master Setup")
'@ModuleDescription("Ribbon callbacks supporting master setup operations.")
'@depends MasterSetupPreparation, IMasterSetupPreparation, MasterSetupHelpers, IDropdownLists, Passwords, IPasswords, Translation, ITranslationObject, TranslationChunks, ITranslationChunks, ApplicationState, IApplicationState, SetupTranslationsTable, ISetupTranslationsTable, IUpdatedValues, DiseaseSheetBuilder, IDiseaseSheetBuilder
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const RIBBON_TRANSLATION_SHEET As String = "__ribbonTranslation"
Private Const RIBBON_TRANSLATION_TABLE As String = "TabTransId"
Private Const RIBBON_LANGUAGE_RANGE As String = "RNG_FileLang"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const PASSWORD_SHEET_NAME As String = "__pass"
Private Const DEFAULT_REMOVE_ROW_TARGET As Long = 1
Private Const DEFAULT_ADD_ROW_BATCH As Long = 10

Private prepService As IMasterSetupPreparation

'@section Manage group callbacks
'===============================================================================
'@Description("Add default rows to tables on the active worksheet.")
'@EntryPoint
Public Sub clickAddRows(ByRef control As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ManageRows targetSheet, True
    End If
End Sub

'@Description("Trim table rows on the active worksheet, preserving header rows.")
'@EntryPoint
Public Sub clickResize(ByRef control As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ManageRows targetSheet, False
    End If
End Sub

'@Description("Clear all active filters applied to tables on the active worksheet.")
'@EntryPoint
Public Sub clickFilters(ByRef control As IRibbonControl)
    Set targetSheet = ActiveSheet
    MasterSetupHelpers.ClearMasterSheetFilters targetSheet
End Sub

'@Description("Sort master setup tables on the active worksheet using default ordering.")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef control As IRibbonControl)
    Dim scope As IApplicationState
    Dim targetSheet As Worksheet

    On Error GoTo Handler

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupHelpers.SortMasterVariablesTables targetSheet

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickRibbonSortTable: "; Err.Number; Err.Description
    Resume Cleanup
End Sub


'@section Disease group callbacks
'===============================================================================
'@Description("Create a new disease worksheet using the builder template.")
'@EntryPoint
Public Sub clickAddSheet(ByRef control As IRibbonControl)
    Dim scope As IApplicationState
    Dim passwords As IPasswords
    Dim dropdowns As IDropdownLists
    Dim translations As ITranslationObject
    Dim builder As IDiseaseSheetBuilder
    Dim diseaseSheet As Worksheet
    Dim diseaseName As String
    Dim attempt As Long
    Dim confirmPrompt As String
    Dim confirmTitle As String
    Dim promptMessage As String
    Dim promptTitle As String
    Dim errorNamePrompt As String
    Dim languageTag As String

    Set translations = MasterSetupHelpers.ResolveRibbonTranslations()
    confirmPrompt = MasterSetupHelpers.TranslateValue(translations, "askConfirmAddDis", "Add a new disease worksheet?")
    confirmTitle = MasterSetupHelpers.TranslateValue(translations, "askConfirm", "Confirm")
    If MsgBox(confirmPrompt, vbYesNo + vbQuestion, confirmTitle) <> vbYes Then Exit Sub

    promptMessage = MasterSetupHelpers.TranslateValue(translations, "enterDis", "Enter the disease name")
    promptTitle = MasterSetupHelpers.TranslateValue(translations, "enterValue", "Disease")
    errorNamePrompt = MasterSetupHelpers.TranslateValue(translations, "errDisName", "Unable to capture the disease name.")

    For attempt = 1 To 5
        diseaseName = MasterSetupHelpers.CleanMasterSheetName(InputBox(promptMessage, promptTitle))
        If LenB(diseaseName) > 0 Then Exit For
    Next attempt

    If LenB(diseaseName) = 0 Then
        If attempt > 5 Then MsgBox errorNamePrompt, vbCritical + vbOKOnly, confirmTitle
        Exit Sub
    End If

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passwords = MasterSetupHelpers.ResolveMasterPasswords()
    If passwords Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns()
    Set builder = DiseaseSheetBuilder.Create(ThisWorkbook, dropdowns, translations)
    languageTag = MasterSetupHelpers.ResolveRibbonLanguageTag()

    passwords.UnProtectWkb ThisWorkbook
    Set diseaseSheet = builder.Build(diseaseName, MasterSetupHelpers.ResolveNextDiseaseIndex(), languageTag)
    passwords.Protect diseaseSheet.Name
    passwords.ProtectWkb ThisWorkbook

    RefreshDropdownCaches
    MsgBox MasterSetupHelpers.TranslateValue(translations, "done", "Done!"), vbInformation + vbOKOnly, confirmTitle

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddSheet: "; Err.Number; Err.Description
    MsgBox MasterSetupHelpers.TranslateValue(translations, "errDisCreate", "Unable to create the disease worksheet."), vbCritical + vbOKOnly, confirmTitle
    If Not passwords Is Nothing Then
        On Error Resume Next
            passwords.ProtectWkb ThisWorkbook
        On Error GoTo 0
    End If
    Resume Cleanup
End Sub

'@Description("Remove the current disease worksheet after confirmation.")
'@EntryPoint
Public Sub clickRemSheet(ByRef control As IRibbonControl)
    Dim scope As IApplicationState
    Dim passwords As IPasswords
    Dim translations As ITranslationObject
    Dim targetSheet As Worksheet
    Dim confirmPrompt As String
    Dim confirmTitle As String
    Dim notDiseaseMessage As String
    Dim alertsState As Boolean

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    Set translations = MasterSetupHelpers.ResolveRibbonTranslations()
    confirmTitle = MasterSetupHelpers.TranslateValue(translations, "askConfirm", "Confirm")
    notDiseaseMessage = MasterSetupHelpers.TranslateValue(translations, "errDisNotFound", "Select a disease worksheet before removing it.")

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox notDiseaseMessage, vbExclamation + vbOKOnly, confirmTitle
        Exit Sub
    End If

    confirmPrompt = MasterSetupHelpers.TranslateValue(translations, "askConfirmRemDis", "Remove the selected disease worksheet?")
    If MsgBox(confirmPrompt, vbYesNo + vbQuestion, confirmTitle) <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passwords = MasterSetupHelpers.ResolveMasterPasswords()
    If passwords Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickRemSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passwords.UnProtect targetSheet.Name
    passwords.UnProtectWkb ThisWorkbook

    alertsState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    targetSheet.Delete
    Application.DisplayAlerts = alertsState

    passwords.ProtectWkb ThisWorkbook

    RefreshDropdownCaches

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickRemSheet: "; Err.Number; Err.Description
    Application.DisplayAlerts = True
    If Not passwords Is Nothing Then
        On Error Resume Next
            passwords.Protect targetSheet.Name
            passwords.ProtectWkb ThisWorkbook
        On Error GoTo 0
    End If
    MsgBox MasterSetupHelpers.TranslateValue(translations, "errDisRemove", "Unable to remove the selected worksheet."), vbCritical + vbOKOnly, confirmTitle
    Resume Cleanup
End Sub

'@Description("Clear data rows within the active disease worksheet tables.")
'@EntryPoint
Public Sub clickClearSheet(ByRef control As IRibbonControl)
    Dim scope As IApplicationState
    Dim translations As ITranslationObject
    Dim passwords As IPasswords
    Dim targetSheet As Worksheet
    Dim confirmPrompt As String
    Dim confirmTitle As String
    Dim notDiseaseMessage As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    Set translations = MasterSetupHelpers.ResolveRibbonTranslations()
    confirmTitle = MasterSetupHelpers.TranslateValue(translations, "askConfirm", "Confirm")
    notDiseaseMessage = MasterSetupHelpers.TranslateValue(translations, "errDisNotFound", "Select a disease worksheet before clearing it.")

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox notDiseaseMessage, vbExclamation + vbOKOnly, confirmTitle
        Exit Sub
    End If

    confirmPrompt = MasterSetupHelpers.TranslateValue(translations, "askConfirmClearDis", "Clear all data in the current disease worksheet?")
    If MsgBox(confirmPrompt, vbYesNo + vbQuestion, confirmTitle) <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passwords = MasterSetupHelpers.ResolveMasterPasswords()
    If passwords Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickClearSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passwords.UnProtect targetSheet.Name
    MasterSetupHelpers.ClearMasterSheetData targetSheet
    passwords.Protect targetSheet.Name

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickClearSheet: "; Err.Number; Err.Description
    MsgBox MasterSetupHelpers.TranslateValue(translations, "errClearDis", "Unable to clear the disease worksheet."), vbCritical + vbOKOnly, confirmTitle
    If Not passwords Is Nothing Then
        On Error Resume Next
            passwords.Protect targetSheet.Name
        On Error GoTo 0
    End If
    Resume Cleanup
End Sub


'@section Translation group callbacks
'===============================================================================
'@Description("Synchronise the translations table with the registry entries.")
'@EntryPoint
Public Sub clickAddTrans(ByRef control As IRibbonControl)
    Dim scope As IApplicationState
    Dim translationsSheet As Worksheet
    Dim translationsTable As ListObject
    Dim registrySheet As Worksheet
    Dim manager As ISetupTranslationsTable
    Dim updatedValues As IUpdatedValues
    Dim passwords As IPasswords
    Dim confirmTitle As String

    confirmTitle = "Translations"
    If MsgBox("Do you want to update the translation sheet?", vbYesNo + vbQuestion, confirmTitle) <> vbYes Then Exit Sub

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then
        MsgBox "Translations sheet was not found.", vbExclamation + vbOKOnly, confirmTitle
        Exit Sub
    End If

    On Error Resume Next
        Set translationsTable = translationsSheet.ListObjects(TRANSLATIONS_TABLE_NAME)
    On Error GoTo 0
    If translationsTable Is Nothing Then
        MsgBox "Translations table was not found.", vbExclamation + vbOKOnly, confirmTitle
        Exit Sub
    End If

    Set registrySheet = MasterSetupHelpers.ResolveMasterRegistrySheet()
    If registrySheet Is Nothing Then
        MsgBox "Registry sheet was not found.", vbExclamation + vbOKOnly, confirmTitle
        Exit Sub
    End If

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passwords = MasterSetupHelpers.ResolveMasterPasswords()
    If passwords Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddTrans", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passwords.UnProtect translationsSheet.Name

    On Error Resume Next
        translationsTable.AutoFilter.ShowAllData
    On Error GoTo 0

    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.UpdateFromRegistry registrySheet

    passwords.Protect translationsSheet.Name

    Set updatedValues = MasterSetupHelpers.ResolveMasterUpdatedValues()
    If Not updatedValues Is Nothing Then updatedValues.SwitchTagsToNo

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    MsgBox "An error occurred while updating translations.", vbCritical + vbOKOnly, confirmTitle
    If Not passwords Is Nothing Then passwords.Protect translationsSheet.Name
    Resume Cleanup
End Sub

'@Description("Add a new language column to the translations table.")
'@EntryPoint
Public Sub clickAddLang(ByRef control As IRibbonControl, ByRef text As String)
    Dim scope As IApplicationState
    Dim workbook As Workbook
    Dim translationsSheet As Worksheet
    Dim translationTagSheet As Worksheet
    Dim translationTable As ListObject
    Dim dropdowns As IDropdownLists
    Dim chunks As ITranslationChunks
    Dim passwords As IPasswords
    Dim translations As ITranslationObject
    Dim languagePrompt As String
    Dim confirmTitle As String
    Dim confirmPrompt As String
    Dim fileLanguage As String

    text = Trim$(text)
    If LenB(text) = 0 Then Exit Sub

    Set workbook = ThisWorkbook

    On Error Resume Next
        Set translationsSheet = workbook.Worksheets(TRANSLATIONS_SHEET_NAME)
        Set translationTagSheet = workbook.Worksheets(RIBBON_TRANSLATION_SHEET)
        Set translationTable = translationTagSheet.ListObjects(RIBBON_TRANSLATION_TABLE)
    On Error GoTo 0

    If translationsSheet Is Nothing Or translationTable Is Nothing Then Exit Sub

    fileLanguage = MasterSetupHelpers.SafeValue(translationTagSheet.Range(RIBBON_LANGUAGE_RANGE).Value)
    Set translations = Translation.Create(translationTable, fileLanguage)
    confirmTitle = MasterSetupHelpers.TranslateValue(translations, "askConfirm", "Confirm")
    confirmPrompt = MasterSetupHelpers.TranslateValue(translations, "addLang", "Add language: ") & text

    If MsgBox(confirmPrompt, vbYesNo + vbQuestion, confirmTitle) <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns()
    Set chunks = TranslationChunks.Create(translationsSheet, TRANSLATIONS_TABLE_NAME, dropdowns)
    Set passwords = MasterSetupHelpers.ResolveMasterPasswords()
    If passwords Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddLang", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passwords.UnProtect TRANSLATIONS_SHEET_NAME
    chunks.AddTransLang text
    passwords.Protect TRANSLATIONS_SHEET_NAME

    languagePrompt = MasterSetupHelpers.TranslateValue(translations, "done", "Done!")
    MsgBox languagePrompt, vbInformation + vbOKOnly, confirmTitle

    RefreshDropdownCaches

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddLang: "; Err.Number; Err.Description
    MsgBox MasterSetupHelpers.TranslateValue(translations, "errAddLang", "Unable to add the language column."), vbCritical + vbOKOnly, confirmTitle
    If Not passwords Is Nothing Then passwords.Protect TRANSLATIONS_SHEET_NAME
    Resume Cleanup
End Sub

'@Description("Change the current ribbon language and refresh workbook labels.")
'@EntryPoint
Public Sub clickLangChange(ByRef control As IRibbonControl, ByVal langId As String, ByVal index As Integer)
    Dim scope As IApplicationState
    Dim tagSheet As Worksheet
    Dim ribbon As IRibbonUI

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set tagSheet = ThisWorkbook.Worksheets(RIBBON_TRANSLATION_SHEET)
    tagSheet.Range(RIBBON_LANGUAGE_RANGE).Value = langId

    RefreshDropdownCaches

    Set ribbon = RibbonDev.ActualRibbon
    If Not ribbon Is Nothing Then ribbon.Invalidate

    On Error Resume Next
        Misc.TranslateWbElmts langId
    On Error GoTo 0

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickLangChange: "; Err.Number; Err.Description
    Resume Cleanup
End Sub

'@Description("Return the translated label for ribbon controls.")
'@EntryPoint
Public Sub LangLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    Dim translations As ITranslationObject
    Dim fallback As String

    On Error GoTo Handler

    Set translations = MasterSetupHelpers.ResolveRibbonTranslations()
    fallback = control.Id

    returnedVal = MasterSetupHelpers.TranslateValue(translations, control.Id, fallback)
    Exit Sub

Handler:
    returnedVal = control.Id
End Sub


'@section Advanced group callbacks
'===============================================================================
'@Description("Export the current disease worksheet to a standalone setup workbook.")
'@EntryPoint
Public Sub clickExpSheet(ByRef control As IRibbonControl)
    Dim scope As IApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Exports.ExportToSetup

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickExpSheet: "; Err.Number; Err.Description
    MsgBox "Disease export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    Resume Cleanup
End Sub

'@Description("Export diseases for migration workflows.")
'@EntryPoint
Public Sub clickExp(ByRef control As IRibbonControl)
    Dim scope As IApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Exports.ExportForMigration

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickExp: "; Err.Number; Err.Description
    MsgBox "Migration export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    Resume Cleanup
End Sub

'@Description("Import diseases from a flat file using the legacy importer.")
'@EntryPoint
Public Sub clickImp(ByRef control As IRibbonControl)
    Dim scope As IApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Exports.ImportFlatFile

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickImp: "; Err.Number; Err.Description
    MsgBox "Disease import failed: " & Err.Description, vbCritical + vbOKOnly, "Import"
    Resume Cleanup
End Sub


'@section Helpers
'===============================================================================
Private Function Preparation() As IMasterSetupPreparation
    If prepService Is Nothing Then
        Set prepService = MasterSetupPreparation.Create(ThisWorkbook)
    End If
    Set Preparation = prepService
End Function

Private Sub RefreshDropdownCaches()
    On Error Resume Next
        Preparation.EnsureDropdowns
    On Error GoTo 0
End Sub
