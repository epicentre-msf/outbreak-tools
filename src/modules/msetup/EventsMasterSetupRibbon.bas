Attribute VB_Name = "EventsMasterSetupRibbon"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Ribbon callbacks supporting master setup operations.")
'@depends MasterSetupPreparation, MasterSetupHelpers, MasterSetupExports, DropdownLists, Passwords, TranslationObject, ApplicationState, SetupTranslationsTable, UpdatedValues, DiseaseSheet, Development, RibbonDev
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'The master setup file itself stays in English: every prompt below is a plain
'English string, and the ribbon carries no translation dropdown. The only
'translatable content is the values inside a disease worksheet.

Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const PASSWORD_SHEET_NAME As String = "__pass"
Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const COMPARE_REPORT_SHEET_NAME As String = "__compRep"
Private Const IMPORT_REPORT_SHEET_NAME As String = "__impRep"
Private Const DEVELOPMENT_SHEET_NAME As String = "Dev"
Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const DEV_PROMPT_TITLE As String = "Development"

Private prepService As MasterSetupPreparation

'@section Ribbon lifecycle
'===============================================================================
'@Description("Cache the ribbon reference when the workbook loads it.")
'@EntryPoint
Public Sub ribbonLoaded(ByRef ribbon As IRibbonUI)
    MasterSetupEventsManager.RibbonLoaded ribbon
End Sub

'@section Manage group callbacks
'===============================================================================
'@Description("Add default rows to tables on the active worksheet.")
'@EntryPoint
Public Sub clickAddRows(ByRef ribbonControl As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ManageRows targetSheet, True
    End If
End Sub

'@Description("Trim table rows on the active worksheet, preserving header rows.")
'@EntryPoint
Public Sub clickResize(ByRef ribbonControl As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ManageRows targetSheet, False
    End If
End Sub

'@Description("Clear all active filters applied to tables on the active worksheet.")
'@EntryPoint
Public Sub clickFilters(ByRef ribbonControl As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ClearMasterSheetFilters targetSheet
    End If
End Sub

'@Description("Sort master setup tables on the active worksheet using default ordering.")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim targetSheet As Worksheet

    On Error GoTo Handler

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupHelpers.SortMasterVariablesTables targetSheet

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickRibbonSortTable: "; Err.Number; Err.Description
    Resume Cleanup
End Sub


'@section Disease group callbacks
'===============================================================================
'@Description("Create a new disease worksheet using the sheet builder.")
'@EntryPoint
Public Sub clickAddSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim dropdowns As DropdownLists
    Dim builder As DiseaseSheet
    Dim diseaseWksh As Worksheet
    Dim diseaseName As String
    Dim attempt As Long

    If MsgBox("Add a new disease worksheet?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    For attempt = 1 To 5
        diseaseName = MasterSetupHelpers.CleanMasterSheetName(InputBox("Enter the disease name", "Disease"))
        If LenB(diseaseName) > 0 Then Exit For
    Next attempt

    If LenB(diseaseName) = 0 Then
        If attempt > 5 Then MsgBox "Unable to capture the disease name.", vbCritical + vbOKOnly, "Confirm"
        Exit Sub
    End If

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns()
    Set builder = DiseaseSheet.Create(ThisWorkbook, dropdowns, _
                                      MasterSetupHelpers.ResolveMasterSetupVariables())

    passManager.UnProtect ThisWorkbook
    'The language of the sheet starts on the first language of the list; the
    'user picks another one on the sheet itself.
    Set diseaseWksh = builder.Build(diseaseName)
    passManager.Protect diseaseWksh.Name
    passManager.Protect ThisWorkbook

    RefreshDropdownCaches
    MsgBox "Done!", vbInformation + vbOKOnly, "Confirm"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddSheet: "; Err.Number; Err.Description
    MsgBox "Unable to create the disease worksheet.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect ThisWorkbook
        On Error GoTo 0
    End If
    Resume Cleanup
End Sub

'@Description("Remove the current disease worksheet after confirmation.")
'@EntryPoint
Public Sub clickRemSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim targetSheet As Worksheet
    Dim alertsState As Boolean

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox "Select a disease worksheet before removing it.", vbExclamation + vbOKOnly, "Confirm"
        Exit Sub
    End If

    If MsgBox("Remove the selected disease worksheet?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickRemSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passManager.UnProtect targetSheet.Name
    passManager.UnProtect ThisWorkbook

    alertsState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    targetSheet.Delete
    Application.DisplayAlerts = alertsState

    passManager.Protect ThisWorkbook

    RefreshDropdownCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickRemSheet: "; Err.Number; Err.Description
    Application.DisplayAlerts = True
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect targetSheet.Name
            passManager.Protect ThisWorkbook
        On Error GoTo 0
    End If
    MsgBox "Unable to remove the selected worksheet.", vbCritical + vbOKOnly, "Confirm"
    Resume Cleanup
End Sub

'@Description("Clear data rows within the active disease worksheet tables.")
'@EntryPoint
Public Sub clickClearSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox "Select a disease worksheet before clearing it.", vbExclamation + vbOKOnly, "Confirm"
        Exit Sub
    End If

    If MsgBox("Clear all data in the current disease worksheet?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickClearSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passManager.UnProtect targetSheet.Name
    MasterSetupHelpers.ClearMasterSheetData targetSheet
    passManager.Protect targetSheet.Name

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickClearSheet: "; Err.Number; Err.Description
    MsgBox "Unable to clear the disease worksheet.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect targetSheet.Name
        On Error GoTo 0
    End If
    Resume Cleanup
End Sub


'@section Translation group callbacks
'===============================================================================
'@Description("Synchronise the translations table with the registry entries.")
'@EntryPoint
Public Sub clickAddTrans(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim translationsSheet As Worksheet
    Dim translationsTable As ListObject
    Dim registrySheet As Worksheet
    Dim manager As SetupTranslationsTable
    'Named updater, because a local called updatedValues would take over the
    'identifier and put the predeclared UpdatedValues instance out of reach in
    'this procedure.
    Dim updater As UpdatedValues
    Dim passManager As Passwords
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

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddTrans", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passManager.UnProtect translationsSheet.Name

    On Error Resume Next
        translationsTable.AutoFilter.ShowAllData
    On Error GoTo 0

    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.UpdateFromRegistry registrySheet

    passManager.Protect translationsSheet.Name

    Set updater = MasterSetupHelpers.ResolveMasterUpdatedValues()
    If Not updater Is Nothing Then updater.SwitchTagsToNo

    'The worksheet functions read the translations; their caches are stale now.
    ResetMasterSetupFunctionCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    MsgBox "An error occurred while updating translations.", vbCritical + vbOKOnly, confirmTitle
    If Not passManager Is Nothing Then passManager.Protect translationsSheet.Name
    Resume Cleanup
End Sub

'@Description("Add a new language column to the translations table.")
'@EntryPoint
Public Sub clickAddLang(ByRef ribbonControl As IRibbonControl, ByRef text As String)
    Dim scope As ApplicationState
    Dim translationsSheet As Worksheet
    Dim translationsTable As ListObject
    Dim manager As SetupTranslationsTable
    Dim passManager As Passwords

    text = Trim$(text)
    If LenB(text) = 0 Then Exit Sub

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set translationsTable = translationsSheet.ListObjects(TRANSLATIONS_TABLE_NAME)
    On Error GoTo 0
    If translationsTable Is Nothing Then Exit Sub

    If MsgBox("Add language: " & text, vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickAddLang", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passManager.UnProtect TRANSLATIONS_SHEET_NAME
    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.EnsureLanguages text
    passManager.Protect TRANSLATIONS_SHEET_NAME

    MsgBox "Done!", vbInformation + vbOKOnly, "Confirm"

    RefreshDropdownCaches
    ResetMasterSetupFunctionCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickAddLang: "; Err.Number; Err.Description
    MsgBox "Unable to add the language column.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then passManager.Protect TRANSLATIONS_SHEET_NAME
    Resume Cleanup
End Sub


'@Description("Delete one language column from the translations table.")
'@EntryPoint
Public Sub clickRemLang(ByRef ribbonControl As IRibbonControl, ByRef text As String)
    Dim scope As ApplicationState
    Dim translationsSheet As Worksheet
    Dim translationsTable As ListObject
    Dim manager As SetupTranslationsTable
    Dim passManager As Passwords

    text = Trim$(text)
    If LenB(text) = 0 Then Exit Sub

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set translationsTable = translationsSheet.ListObjects(TRANSLATIONS_TABLE_NAME)
    On Error GoTo 0
    If translationsTable Is Nothing Then Exit Sub

    If MsgBox("Remove language: " & text, vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickRemLang", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    passManager.UnProtect TRANSLATIONS_SHEET_NAME
    Set manager = SetupTranslationsTable.Create(translationsTable)
    manager.RemoveLanguage text
    passManager.Protect TRANSLATIONS_SHEET_NAME

    MsgBox "Done!", vbInformation + vbOKOnly, "Confirm"

    RefreshDropdownCaches
    ResetMasterSetupFunctionCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickRemLang: "; Err.Number; Err.Description
    MsgBox "Unable to remove the language column: " & Err.Description, vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then passManager.Protect TRANSLATIONS_SHEET_NAME
    Resume Cleanup
End Sub

'@Description("Import passwords from an external workbook, like in the designer.")
'@EntryPoint
Public Sub clickMsImpPass(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim importBook As Workbook
    Dim importer As Passwords
    Dim target As Passwords
    Dim scope As ApplicationState
    Dim passSheet As Worksheet

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passSheet = MasterSetupHelpers.ResolveMasterPasswordsSheet()
    If passSheet Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickMsImpPass", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set importer = Passwords.Create(importBook.Worksheets(1))
    Set target = Passwords.Create(passSheet)
    target.ImportFrom importer

    MsgBox "Done!", vbInformation + vbOKOnly, "Passwords"

Cleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Not scope Is Nothing Then scope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickMsImpPass: "; Err.Number; Err.Description
        MsgBox "Unable to import passwords: " & Err.Description, vbExclamation + vbOKOnly, "Passwords"
        Err.Clear
    End If
End Sub


'@section Advanced group callbacks
'===============================================================================
'@Description("Export the current disease worksheet to a standalone setup workbook.")
'@EntryPoint
Public Sub clickExpSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupExports.ExportToSetup

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickExpSheet: "; Err.Number; Err.Description
    MsgBox "Disease export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    Resume Cleanup
End Sub

'@Description("Export every disease for migration workflows.")
'@EntryPoint
Public Sub clickExp(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupExports.ExportForMigration

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickExp: "; Err.Number; Err.Description
    MsgBox "Migration export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    Resume Cleanup
End Sub

'@Description("Import diseases from a flat migration file.")
'@EntryPoint
Public Sub clickImp(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, blockSecurity:=True

    MasterSetupExports.ImportFlatFile

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "clickImp: "; Err.Number; Err.Description
    MsgBox "Disease import failed: " & Err.Description, vbCritical + vbOKOnly, "Import"
    Resume Cleanup
End Sub


'@section Dev group callbacks
'===============================================================================
'The rest of the Dev group is served by RibbonDev, which every product shares.
'Initialize sits here because each product prepares itself its own way: the
'master setup rebuilds its dropdowns and its variables table, then hands the
'Development manager the sheets Deploy has to hide and to protect.

'@Description("Prepare the master setup workbook for deployment.")
'@EntryPoint
Public Sub clickDevInitialize(ByRef ribbonControl As IRibbonControl)
    Dim manager As Development

    On Error GoTo Handler

    'RibbonDev holds one Development manager for the whole VBA session and the
    'Deploy button reads that same object, so the sheet lists are written on it.
    Set manager = RibbonDev.EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    'Prepare opens and restores its own busy scope.
    Preparation.Prepare Application
    RegisterDeploymentSheets manager

    MsgBox "Done!", vbInformation + vbOKOnly, DEV_PROMPT_TITLE
    Exit Sub

Handler:
    Debug.Print "clickDevInitialize: "; Err.Number; Err.Description
    MsgBox "Initialisation failed: " & Err.Description, vbCritical + vbOKOnly, DEV_PROMPT_TITLE
End Sub


'@section Helpers
'===============================================================================
Private Function Preparation() As MasterSetupPreparation
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

'Deploy sets every sheet handed to AddHiddenSheet very hidden and protects
'every sheet handed to AddProtectedSheet. Both lists are written again on each
'initialise; Development drops a name it already holds, so a second press
'leaves one entry per sheet.
'
'Disease worksheets stay out of both lists. They hold the content the user
'edits after the file ships.
Private Sub RegisterDeploymentSheets(ByVal manager As Development)
    HideOnDeploy manager, DROPDOWNS_SHEET_NAME
    HideOnDeploy manager, REGISTRY_SHEET_NAME
    HideOnDeploy manager, PASSWORD_SHEET_NAME
    HideOnDeploy manager, COMPARE_REPORT_SHEET_NAME
    HideOnDeploy manager, IMPORT_REPORT_SHEET_NAME
    HideOnDeploy manager, DEVELOPMENT_SHEET_NAME

    'Variables and Choices keep row deletion, the same permission
    'MasterSetupHelpers.ProtectMasterSetupSheet gives those two sheets.
    ProtectOnDeploy manager, VARIABLES_SHEET_NAME, True, True
    ProtectOnDeploy manager, CHOICES_SHEET_NAME, True, True
    ProtectOnDeploy manager, TRANSLATIONS_SHEET_NAME, False, False
End Sub

'AddHiddenSheet and AddProtectedSheet raise ElementNotFound on a name the
'workbook has no sheet for, so every name is looked up first. A master setup
'that carries only some of the report sheets still initialises.
Private Sub HideOnDeploy(ByVal manager As Development, ByVal sheetName As String)
    If DeploymentSheet(sheetName) Is Nothing Then Exit Sub
    manager.AddHiddenSheet sheetName
End Sub

Private Sub ProtectOnDeploy(ByVal manager As Development, _
                            ByVal sheetName As String, _
                            ByVal allowShapes As Boolean, _
                            ByVal allowDeletingRows As Boolean)
    If DeploymentSheet(sheetName) Is Nothing Then Exit Sub
    manager.AddProtectedSheet sheetName, allowShapes, allowDeletingRows
End Sub

Private Function DeploymentSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set DeploymentSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function
