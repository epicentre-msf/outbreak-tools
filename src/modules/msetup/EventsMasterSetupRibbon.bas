Attribute VB_Name = "EventsMasterSetupRibbon"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Ribbon callbacks supporting master setup operations.")
'@depends MasterSetupEventsManager, EventMasterSetup, MasterSetupLog, MasterSetupPreparation, MasterSetupHelpers, MasterSetupExports, DiseaseSheet, DropdownLists, SetupTranslationsTable, UpdatedValues, Development, RibbonDev, Passwords, ApplicationState, OSFiles
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'The master setup file itself stays in English: every prompt below is a plain
'English string, and the ribbon carries no translation dropdown. The only
'translatable content is the values inside a disease worksheet.
'
'Every callback writes its outcome in the user log EventMasterSetup holds
'on __log: one success line when the walk ends well, one warning line when
'the user cancels or the walk is refused, one failure line at the error
'label. The doors of MasterSetupExports write their own success and warning
'lines; the callbacks that open them write the failure line, because the
'raise lands here. Every write is guarded so a log fault never takes down
'the walk it records.

Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const PASSWORD_SHEET_NAME As String = "__pass"
Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const COMPARE_REPORT_SHEET_NAME As String = "__compRep"
Private Const IMPORT_REPORT_SHEET_NAME As String = "__impRep"
Private Const LOG_SHEET_NAME As String = "__log"
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
    Dim failureText As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    'ManageRows restores the application state and the protection itself
    'before it raises, so the failure only has to be said here.
    On Error GoTo Handler
    MasterSetupHelpers.ManageRows targetSheet, True
    LogSuccessLine "add-rows", targetSheet.Name, "clickAddRows"
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickAddRows: "; Err.Number; Err.Description
    MsgBox "Unable to add rows: " & Err.Description, vbCritical + vbOKOnly, "Rows"
    LogFailureLine "add-rows", failureText, "clickAddRows"
End Sub

'@Description("Trim table rows on the active worksheet, preserving header rows.")
'@EntryPoint
Public Sub clickResize(ByRef ribbonControl As IRibbonControl)
    Dim targetSheet As Worksheet
    Dim failureText As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    On Error GoTo Handler
    MasterSetupHelpers.ManageRows targetSheet, False
    LogSuccessLine "resize-tables", targetSheet.Name, "clickResize"
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickResize: "; Err.Number; Err.Description
    MsgBox "Unable to resize the tables: " & Err.Description, vbCritical + vbOKOnly, "Rows"
    LogFailureLine "resize-tables", failureText, "clickResize"
End Sub

'@Description("Clear all active filters applied to tables on the active worksheet.")
'@EntryPoint
Public Sub clickFilters(ByRef ribbonControl As IRibbonControl)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    If Not targetSheet Is Nothing Then
        MasterSetupHelpers.ClearMasterSheetFilters targetSheet
        LogSuccessLine "clear-filters", targetSheet.Name, "clickFilters"
    End If
End Sub

'@Description("Sort master setup tables on the active worksheet using default ordering.")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim targetSheet As Worksheet
    Dim failureText As String

    On Error GoTo Handler

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupHelpers.SortMasterVariablesTables targetSheet
    LogSuccessLine "sort-tables", targetSheet.Name, "clickRibbonSortTable"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickRibbonSortTable: "; Err.Number; Err.Description
    LogFailureLine "sort-tables", failureText, "clickRibbonSortTable"
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
    Dim failureText As String

    If MsgBox("Add a new disease worksheet?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    For attempt = 1 To 5
        diseaseName = MasterSetupHelpers.CleanMasterSheetName(InputBox("Enter the disease name", "Disease"))
        If LenB(diseaseName) > 0 Then Exit For
    Next attempt

    If LenB(diseaseName) = 0 Then
        If attempt > 5 Then MsgBox "Unable to capture the disease name.", vbCritical + vbOKOnly, "Confirm"
        LogWarningLine "add-disease", "no disease name given", "clickAddSheet"
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
    LogSuccessLine "add-disease", diseaseName, "clickAddSheet"
    MsgBox "Done!", vbInformation + vbOKOnly, "Confirm"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickAddSheet: "; Err.Number; Err.Description
    MsgBox "Unable to create the disease worksheet.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect ThisWorkbook
        On Error GoTo 0
    End If
    LogFailureLine "add-disease", diseaseName & ": " & failureText, "clickAddSheet"
    Resume Cleanup
End Sub

'@Description("Remove the current disease worksheet after confirmation.")
'@EntryPoint
Public Sub clickRemSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim targetSheet As Worksheet
    Dim alertsState As Boolean
    Dim removedName As String
    Dim failureText As String

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

    removedName = targetSheet.Name
    passManager.UnProtect targetSheet.Name
    passManager.UnProtect ThisWorkbook

    alertsState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    targetSheet.Delete
    Application.DisplayAlerts = alertsState

    passManager.Protect ThisWorkbook

    RefreshDropdownCaches
    LogSuccessLine "remove-disease", removedName, "clickRemSheet"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickRemSheet: "; Err.Number; Err.Description
    Application.DisplayAlerts = True
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect targetSheet.Name
            passManager.Protect ThisWorkbook
        On Error GoTo 0
    End If
    MsgBox "Unable to remove the selected worksheet.", vbCritical + vbOKOnly, "Confirm"
    LogFailureLine "remove-disease", removedName & ": " & failureText, "clickRemSheet"
    Resume Cleanup
End Sub

'@Description("Clear data rows within the active disease worksheet tables.")
'@EntryPoint
Public Sub clickClearSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim targetSheet As Worksheet
    Dim failureText As String

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
    LogSuccessLine "clear-disease", targetSheet.Name, "clickClearSheet"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickClearSheet: "; Err.Number; Err.Description
    MsgBox "Unable to clear the disease worksheet.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect targetSheet.Name
        On Error GoTo 0
    End If
    LogFailureLine "clear-disease", targetSheet.Name & ": " & failureText, "clickClearSheet"
    Resume Cleanup
End Sub


'@Description("Compare two disease worksheets and open the report on __compRep.")
'@EntryPoint
Public Sub clickCompare(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim failureText As String

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    MasterSetupExports.CompareDiseaseSheets

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickCompare: "; Err.Number; Err.Description
    MsgBox "Disease comparison failed: " & Err.Description, vbCritical + vbOKOnly, "Compare"
    LogFailureLine "compare-diseases", failureText, "clickCompare"
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
    Dim failureText As String

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
    LogSuccessLine "update-translations", vbNullString, "clickAddTrans"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickAddTrans: "; Err.Number; Err.Description
    MsgBox "An error occurred while updating translations.", vbCritical + vbOKOnly, confirmTitle
    If Not passManager Is Nothing Then passManager.Protect translationsSheet.Name
    LogFailureLine "update-translations", failureText, "clickAddTrans"
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
    Dim failureText As String

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

    LogSuccessLine "add-language", text, "clickAddLang"
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
    failureText = Err.Description
    Debug.Print "clickAddLang: "; Err.Number; Err.Description
    MsgBox "Unable to add the language column.", vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then passManager.Protect TRANSLATIONS_SHEET_NAME
    LogFailureLine "add-language", text & ": " & failureText, "clickAddLang"
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
    Dim failureText As String

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

    LogSuccessLine "remove-language", text, "clickRemLang"
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
    failureText = Err.Description
    Debug.Print "clickRemLang: "; Err.Number; Err.Description
    MsgBox "Unable to remove the language column: " & Err.Description, vbCritical + vbOKOnly, "Confirm"
    If Not passManager Is Nothing Then passManager.Protect TRANSLATIONS_SHEET_NAME
    LogFailureLine "remove-language", text & ": " & failureText, "clickRemLang"
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
    Dim failureText As String

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then
        'The cancel is worth a line: a user who says the import did nothing
        'has the record of their own cancel.
        LogWarningLine "import-passwords", "file picker cancelled", "clickMsImpPass"
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set passSheet = MasterSetupHelpers.ResolveMasterPasswordsSheet()
    If passSheet Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickMsImpPass", "Passwords sheet '" & PASSWORD_SHEET_NAME & "' was not found."

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set importer = Passwords.Create(importBook.Worksheets(1))
    Set target = Passwords.Create(passSheet)
    target.ImportFrom importer

    LogSuccessLine "import-passwords", io.File(), "clickMsImpPass"
    MsgBox "Done!", vbInformation + vbOKOnly, "Passwords"

Cleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Not scope Is Nothing Then scope.Restore
    If Err.Number <> 0 Then
        failureText = Err.Description
        Debug.Print "clickMsImpPass: "; Err.Number; Err.Description
        MsgBox "Unable to import passwords: " & Err.Description, vbExclamation + vbOKOnly, "Passwords"
        Err.Clear
        LogFailureLine "import-passwords", failureText, "clickMsImpPass"
    End If
End Sub


'@section Advanced group callbacks
'===============================================================================
'@Description("Export the current disease worksheet to a standalone setup workbook.")
'@EntryPoint
Public Sub clickExpSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim failureText As String

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
    failureText = Err.Description
    Debug.Print "clickExpSheet: "; Err.Number; Err.Description
    MsgBox "Disease export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    LogFailureLine "export-setup", failureText, "clickExpSheet"
    Resume Cleanup
End Sub

'@Description("Import a disease workbook exported for a setup back into this file.")
'@EntryPoint
Public Sub clickImpSheet(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim failureText As String

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, blockSecurity:=True

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickImpSheet", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    'The import writes on the Variables and Choices sheets, on the disease
    'sheet it lands on, and adds a sheet when the disease is new.
    SetMasterSheetsProtection passManager, protectSheets:=False
    MasterSetupExports.ImportFromSetup
    SetMasterSheetsProtection passManager, protectSheets:=True

    RefreshDropdownCaches
    ResetMasterSetupFunctionCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickImpSheet: "; Err.Number; Err.Description
    MsgBox "Disease import failed: " & Err.Description, vbCritical + vbOKOnly, "Import"
    If Not passManager Is Nothing Then
        On Error Resume Next
            SetMasterSheetsProtection passManager, protectSheets:=True
        On Error GoTo 0
    End If
    LogFailureLine "import-setup", failureText, "clickImpSheet"
    Resume Cleanup
End Sub

'@Description("Export the whole master setup into one migration file.")
'@EntryPoint
Public Sub clickExp(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim failureText As String

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
    failureText = Err.Description
    Debug.Print "clickExp: "; Err.Number; Err.Description
    MsgBox "Migration export failed: " & Err.Description, vbCritical + vbOKOnly, "Export"
    LogFailureLine "export-migration", failureText, "clickExp"
    Resume Cleanup
End Sub

'@Description("Import the whole master setup from a migration file into this empty file.")
'@EntryPoint
Public Sub clickImp(ByRef ribbonControl As IRibbonControl)
    Dim scope As ApplicationState
    Dim passManager As Passwords
    Dim failureText As String

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, blockSecurity:=True

    Set passManager = MasterSetupHelpers.ResolveMasterPasswords()
    If passManager Is Nothing Then Err.Raise ProjectError.ElementNotFound, "clickImp", "Passwords worksheet '" & PASSWORD_SHEET_NAME & "' was not found."

    'The migration writes on the Variables, Choices and Translations sheets
    'and adds one worksheet per disease of the file.
    SetMasterSheetsProtection passManager, protectSheets:=False
    passManager.UnProtect TRANSLATIONS_SHEET_NAME
    MasterSetupExports.ImportFlatFile
    passManager.Protect TRANSLATIONS_SHEET_NAME
    SetMasterSheetsProtection passManager, protectSheets:=True

    RefreshDropdownCaches
    ResetMasterSetupFunctionCaches

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    failureText = Err.Description
    Debug.Print "clickImp: "; Err.Number; Err.Description
    MsgBox "Migration import failed: " & Err.Description, vbCritical + vbOKOnly, "Import"
    If Not passManager Is Nothing Then
        On Error Resume Next
            passManager.Protect TRANSLATIONS_SHEET_NAME
            SetMasterSheetsProtection passManager, protectSheets:=True
        On Error GoTo 0
    End If
    LogFailureLine "import-migration", failureText, "clickImp"
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

'@section User log
'===============================================================================
'The user log the event service holds. A workbook whose log cannot be built
'answers Nothing and every log line of this module stays quiet.
Private Function UserLogOf() As MasterSetupLog
    Dim service As EventMasterSetup

    Set service = MasterSetupEventsManager.MasterSetupService()
    If service Is Nothing Then Exit Function

    Set UserLogOf = service.UserLog()
End Function

'Write the success line of a finished walk. The write is guarded so a log
'fault never takes down the walk it records.
Private Sub LogSuccessLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal source As String = vbNullString)
    Dim logStore As MasterSetupLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail, source
    On Error GoTo 0
End Sub

'Write the warning line of a refused or cancelled walk.
Private Sub LogWarningLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal source As String = vbNullString)
    Dim logStore As MasterSetupLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail, source
    On Error GoTo 0
End Sub

'Write the failure line of a walk that ended at its error label. The caller
'copies Err.Description before calling: the guard below resets Err.
Private Sub LogFailureLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal source As String = vbNullString)
    Dim logStore As MasterSetupLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogFailure action, detail, source
    On Error GoTo 0
End Sub

'The workbook, the Variables and Choices sheets and every disease worksheet
'open before a setup import and close after it: the import can write on any
'of them, and the disease sheet it adds is protected with the others.
Private Sub SetMasterSheetsProtection(ByVal passManager As Passwords, ByVal protectSheets As Boolean)
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
        If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
            If protectSheets Then
                passManager.Protect sh.Name
            Else
                passManager.UnProtect sh.Name
            End If
        End If
    Next sh

    If protectSheets Then
        MasterSetupHelpers.ProtectMasterSetupSheet DeploymentSheet(VARIABLES_SHEET_NAME), "variables"
        MasterSetupHelpers.ProtectMasterSetupSheet DeploymentSheet(CHOICES_SHEET_NAME), "choices"
        passManager.Protect ThisWorkbook
    Else
        passManager.UnProtect ThisWorkbook
        MasterSetupHelpers.UnProtectMasterSetupSheet DeploymentSheet(VARIABLES_SHEET_NAME), "variables"
        MasterSetupHelpers.UnProtectMasterSetupSheet DeploymentSheet(CHOICES_SHEET_NAME), "choices"
    End If
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
    HideOnDeploy manager, LOG_SHEET_NAME
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
