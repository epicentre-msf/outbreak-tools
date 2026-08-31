Attribute VB_Name = "SetupRibbon"

Option Explicit

'@Folder("Events")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used
'@depends EventsManager, EventSetup, SetupHelpers, SetupPreparation, SetupImport, SetupTranslationsTable, UpdatedValues, RibbonDev, TranslationObject, Passwords, BetterArray, Messenger

'Every callback here reaches the setup service through
'EventsManager.EventSetupService. Row work, sorting, sheet protection and
'the setup translation all live on EventSetup, so a click crosses the ribbon, the
'manager and the service, and nothing in between.
'
'A callback that asks the user a question asks it BEFORE entering busy state. A
'prompt raised over a frozen screen with a busy cursor reads as a hang.
'
'THE EIGHT BOXES A SCRIPT CAN WALK INTO GO THROUGH THE MESSENGER
'Three boxes in RunTranslationsUpdate, two in clickExport and two in
'clickImportFile call Messenger.Show. The closing summary of the tag update
'does the same, over in SetupTranslationsTable. All eight are vbOKOnly, so
'every silent answer is vbOK. While the messenger is disarmed each one opens
'the box it always opened; while it is armed the text is written down and
'Messenger.Messages reads it back. The other 21 boxes in this module sit on
'buttons the R package never presses.
'
'THREE ENTRY POINTS A SCRIPT CAN CALL SIT AT THE FOOT OF THIS MODULE
'RunSetupExport, RunSetupImportFile and RunSetupTags take their paths as
'strings, open no picker and no box, and answer "OK" or "ERROR <number>:
'<text>". SetupLastSummary reads the last of those runs back. The buttons
'above them are unchanged.

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_EXPORTS As String = "Exports"

'THE R PACKAGE ENTRY POINTS KEEP THEIR STATE HERE
'A module-level declaration has to sit in this section, above every procedure.
'These lived beside the wrappers at the foot of the module until they did not
'compile: VBA registers no declaration written between two procedures, and every
'use of them then reads as an undefined variable under Option Explicit.

'What a wrapper answers when the run went through.
Private Const OUTCOME_OK As String = "OK"

'What an outcome that failed opens with.
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "

'What the summary file is called, after the name of the file the run touched.
Private Const SUMMARY_SUFFIX As String = "-obt-summary.txt"

'The marker the free text starts after, the shape HeadlessBuild already answers.
Private Const REPORT_MARKER As String = "--report--"

'What the last wrapper run wrote, read and said. SetupLastSummary answers these,
'and they are module level because that reading is a second Application.Run.
Private lastOutcome As String
Private lastExportFile As String
Private lastImportFile As String
Private lastMessages As String

'The row buttons, the column button and the sort button are always visible.
'They used to hide themselves through a getVisible callback that the events
'invalidated on each sheet change; now each one checks the active sheet when
'it is clicked and tells the user where it works instead.

'@sub-title Whether the active sheet holds the tables the row buttons work on.
Private Function IsRowSheet(ByVal sheetName As String) As Boolean
    Select Case LCase$(sheetName)
        Case LCase$(SHEET_DICTIONARY), LCase$(SHEET_CHOICES), _
             LCase$(SHEET_ANALYSIS), LCase$(SHEET_EXPORTS)
            IsRowSheet = True
    End Select
End Function

'@sub-title Tell the user a button was clicked on a sheet it does not work on.
Private Sub WarnWrongSheet(ByVal actionName As String, ByVal sheetList As String)
    MsgBox actionName & " works only on the " & sheetList & ".", vbExclamation, actionName
End Sub

'@sub-title Warn when a row button is clicked outside the four table sheets.
Private Function OnRowSheet(ByVal sheetName As String, ByVal actionName As String) As Boolean
    OnRowSheet = IsRowSheet(sheetName)
    If Not OnRowSheet Then WarnWrongSheet actionName, "Dictionary, Choices, Analysis and Exports sheets"
End Function


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
    If Not OnRowSheet(sheetName, "Sort tables") Then Exit Sub
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
    If Not OnRowSheet(sheetName, "Insert row") Then Exit Sub
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
    If Not OnRowSheet(ActiveSheet.Name, "Delete table row") Then Exit Sub

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
    If StrComp(sheetName, TRADSHEETNAME, vbTextCompare) <> 0 Then
        WarnWrongSheet "Delete table column", "Translations sheet"
        Exit Sub
    End If

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
        Messenger.Show "Translations table was not found.", vbOK, vbExclamation
        Exit Function
    End If

    Set registrySheet = SetupHelpers.ResolveRegistrySheet
    If registrySheet Is Nothing Then
        Messenger.Show "Registry sheet was not found.", vbOK, vbExclamation
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
    Messenger.Show "An error occurred while updating translations.", vbOK, vbCritical
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
        Messenger.Show "Setup exported to: " & vbCrLf & exportPath, vbOK, vbInformation
    End If

    Exit Sub

Handler:
    'Re-protect BEFORE restoring screen state to avoid visible flash
    On Error Resume Next
    svc.ProtectSetupSheet SetupHelpers.ResolveSetupSheetName("ana")
    On Error GoTo 0
    EventsManager.ExitBusyState
    Debug.Print "clickExport: "; Err.Number; Err.Description
    Messenger.Show "Failed to export the setup: " & Err.Description, vbOK, vbCritical
End Sub

'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickImport(ByRef ribbonControl As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=False
    [Imports].Show

    'Excel hands the pointer back on the default cursor when a modal form
    'closes, and from there every macro of the session flashes its busy
    'pointer. The manager parks it back on the arrow the setup rests on.
    EventsManager.RestPointer
End Sub


'@Description("Callback for btnImp onAction: import setup content from another setup workbook")
'@EntryPoint
Public Sub clickClearSetup(ByRef ribbonControl As IRibbonControl)
    SetupHelpers.PrepareImportsForm cleanSetup:=True
    [Imports].Show

    'Excel hands the pointer back on the default cursor when a modal form
    'closes, and from there every macro of the session flashes its busy
    'pointer. The manager parks it back on the arrow the setup rests on.
    EventsManager.RestPointer
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
    If success Then Messenger.Show "Import Done!", vbOK
    Exit Sub

Handler:
    Debug.Print "clickImportFile: "; Err.Number; Err.Description
    success = False
    Messenger.Show "Failed to import workbook data: " & Err.Description, vbOK, vbCritical
    Resume Cleanup
End Sub

Public Sub clickCheck(ByRef ribbonControl As IRibbonControl)
    On Error GoTo Cleanup

    EventsManager.EnterBusyState
    SetupHelpers.CheckTheSetup

Cleanup:
    EventsManager.ExitBusyState
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


'@section The R package entry points
'===============================================================================
'THREE WRAPPERS AND ONE SUMMARY, FOR A SCRIPT DRIVING THIS WORKBOOK
'-------------------------------------------------------------------------------
'A ribbon callback takes an IRibbonControl as its first argument, and a script
'outside Excel cannot build one, so Application.Run cannot reach a button. The
'three functions below are what a script calls instead. Each one:
'
'   takes its paths as plain strings, so nothing has to be filled in first,
'   arms the Messenger, so no box waits for a click on any path,
'   checks its file or folder before it starts,
'   runs the body its button runs, and
'   answers "OK" or "ERROR <number>: <text>".
'
'The buttons keep their pickers, their questions and their boxes. A person
'clicking in Excel sees no difference at all.
'
'EVERY PATH DISARMS
'-------------------------------------------------------------------------------
'CloseRun is the one exit of all three: it reads the swallowed messages off the
'messenger, disarms, writes the summary file and answers the outcome. A run that
'refuses at its guard goes through it too, so a bad path cannot leave the next
'caller's boxes swallowed.
'
'THE SUMMARY IS WRITTEN TWICE
'-------------------------------------------------------------------------------
'SetupLastSummary reads the run back through a second Application.Run, and that
'is the reading that is lost whenever the Apple Event transport gives up -- which
'happens here on runs Excel finished green. So the same text is written to
'<folder>/<name>-obt-summary.txt beside the file each wrapper touched, and the
'file survives a -1712 and a wedged Excel both.

'@sub-title Export the whole setup to a workbook, with no picker and no box.
'@details
'The body of clickExport, with the folder handed in instead of picked.
'SetExportFolder and DisplayPrompts False are what keep EnsureExportWorkbook
'away from its folder dialog.
'
'Export REFUSES QUIETLY. It is a Sub, and EnsureExportWorkbook answers Nothing
'when prompts are off with no folder set, on which Export does Exit Sub with no
'error at all. So the outcome is read off LastExportFile: an empty answer is a
'failure here, whatever Export did or did not raise.
'@param outputFolder String. Where the .xlsx goes. Empty means the folder this
'workbook sits in.
'@return String. "OK", or "ERROR <number>: <text>".
'@EntryPoint
Public Function RunSetupExport(ByVal outputFolder As String) As String
    Dim svc As EventSetup
    Dim service As SetupImport
    Dim analysisSheet As String
    Dim targetFolder As String
    Dim outcome As String
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    Messenger.Arm ThisWorkbook
    ResetRunRecord

    targetFolder = Trim$(outputFolder)
    If LenB(targetFolder) = 0 Then targetFolder = ThisWorkbook.Path

    If Not FolderIsThere(targetFolder) Then
        RunSetupExport = CloseRun(RunError(0, "no folder at " & targetFolder), _
                                  ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
        Exit Function
    End If

    Set svc = EventsManager.EventSetupService

    On Error GoTo Handler

    analysisSheet = SetupHelpers.ResolveSetupSheetName("ana")

    EventsManager.EnterBusyState

    Set service = SetupImport.Create(ThisWorkbook.FullName)
    service.DisplayPrompts = False
    service.SetExportFolder targetFolder

    'UnProtect the analysis before proceeding
    svc.UnprotectSetupSheet analysisSheet
    sheetUnlocked = True
    service.Export
    svc.ProtectSetupSheet analysisSheet
    sheetUnlocked = False

    lastExportFile = service.LastExportFile

    EventsManager.ExitBusyState

    If LenB(lastExportFile) = 0 Then
        outcome = RunError(0, "the export wrote no file into " & targetFolder)
    Else
        Messenger.Show "Setup exported to: " & vbCrLf & lastExportFile, vbOK, vbInformation
        outcome = OUTCOME_OK
    End If

    RunSetupExport = CloseRun(outcome, targetFolder, SummaryBase(lastExportFile))
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description

    'Re-protect BEFORE restoring screen state to avoid visible flash
    On Error Resume Next
        If sheetUnlocked Then svc.ProtectSetupSheet analysisSheet
    On Error GoTo 0
    EventsManager.ExitBusyState

    RunSetupExport = CloseRun(RunError(errNumber, errDescription), _
                              targetFolder, SummaryBase(lastExportFile))
End Function

'@sub-title Read another setup workbook into this one, with no picker and no box.
'@details
'The body of clickImportFile, with the path handed in instead of picked. The
'picker also proved the file was there; a path from a script has proved nothing,
'so the file is checked before any of the work starts.
'
'This never enters SetupHelpers.ImportOrCleanSetup, which is the [Imports] form's
'path and holds four boxes of its own.
'@param importPath String. Full path of the .xlsx to read.
'@return String. "OK", or "ERROR <number>: <text>".
'@EntryPoint
Public Function RunSetupImportFile(ByVal importPath As String) As String
    Dim service As SetupImport
    Dim pass As Passwords
    Dim sheets As BetterArray
    Dim originalSheet As Worksheet
    Dim sourcePath As String
    Dim outcome As String
    Dim errNumber As Long
    Dim errDescription As String

    Messenger.Arm ThisWorkbook
    ResetRunRecord

    sourcePath = Trim$(importPath)

    If Not FileIsThere(sourcePath) Then
        RunSetupImportFile = CloseRun(RunError(0, "no file at " & sourcePath), _
                                      ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
        Exit Function
    End If

    On Error GoTo Handler

    Set service = SetupImport.Create(sourcePath)
    service.DisplayPrompts = False
    Set pass = SetupHelpers.ResolveSetupPasswords()
    Set sheets = SetupHelpers.BuildSelectedSheets(True, True, True, True, True)
    Set originalSheet = ActiveSheet

    EventsManager.EnterBusyState calculateOnSave:=False

    service.ImportFromWorkbook pass, sheets
    SetupHelpers.PostImportMaintenance

    lastImportFile = sourcePath
    outcome = OUTCOME_OK
    Messenger.Show "Import Done!", vbOK

Cleanup:
    EventsManager.ExitBusyState
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.ScreenUpdating = True

    RunSetupImportFile = CloseRun(outcome, ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print "RunSetupImportFile: "; errNumber; errDescription
    outcome = RunError(errNumber, errDescription)
    Resume Cleanup
End Function

'@sub-title Reset the update tags and rebuild the translations, with no box.
'@details
'The body of clickResetTag, with the review off. Reset tags runs the update with
'the review ON, which asks about every label no range of the setup produces;
'nobody is there to answer that, so the wrapper runs it OFF and removeUnseen
'decides on its own, exactly as the Update Translations button does.
'@return String. "OK", or "ERROR <number>: <text>".
'@EntryPoint
Public Function RunSetupTags() As String
    Dim prep As SetupPreparation
    Dim outcome As String
    Dim errNumber As Long
    Dim errDescription As String

    Messenger.Arm ThisWorkbook
    ResetRunRecord

    On Error GoTo Handler

    EventsManager.EnterBusyState

    Set prep = SetupPreparation.Create(ThisWorkbook)
    prep.ResetUpdatedRegistry

    EventsManager.ExitBusyState

    If RunTranslationsUpdate(reviewUnseen:=False) Then
        outcome = OUTCOME_OK
    Else
        outcome = RunError(0, "the translations were not updated")
    End If

    RunSetupTags = CloseRun(outcome, ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print "RunSetupTags: "; errNumber; errDescription
    EventsManager.ExitBusyState

    RunSetupTags = CloseRun(RunError(errNumber, errDescription), _
                            ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
End Function

'@sub-title What the last wrapper run wrote, read and said.
'@details
'key=value lines, then the REPORT_MARKER, then the boxes the run swallowed, one
'per line. The shape HeadlessBuild.LastBuildSummary already answers.
'
'outcome= leads the block. The whole reason this text is also written to a file
'is that the answer of Application.Run is lost when the transport gives up, and
'the outcome is the first thing that reading loses, so the file has to carry it.
'
'The setup answers paths and nothing more. Counts belong to the generation run,
'which keeps its own record in the designer's GenerationLog.
'@return String. The summary of the last run, empty keys and all.
'@EntryPoint
Public Function SetupLastSummary() As String
    SetupLastSummary = "outcome=" & lastOutcome & vbLf & _
                       "export=" & lastExportFile & vbLf & _
                       "imported=" & lastImportFile & vbLf & _
                       REPORT_MARKER & vbLf & lastMessages
End Function


'@section What the three wrappers share
'===============================================================================

'@sub-title Forget the run before this one.
'@details
'Messenger.Arm empties its own record; this empties what SetupLastSummary reads
'beside it, so a refused run never answers the paths of the run before it.
Private Sub ResetRunRecord()
    lastOutcome = vbNullString
    lastExportFile = vbNullString
    lastImportFile = vbNullString
    lastMessages = vbNullString
End Sub

'@sub-title The one exit of all three wrappers.
'@details
'Reads the swallowed boxes, disarms the messenger, writes the summary beside the
'file the run touched, and answers the outcome. The file write is deliberately
'quiet: a run that worked is not turned into a failure because the folder it was
'given cannot be written to.
'@param outcome String. "OK", or an "ERROR ..." line.
'@param folderPath String. Where the summary file goes.
'@param baseName String. What the summary file is named after.
'@return String. The outcome it was given.
Private Function CloseRun(ByVal outcome As String, _
                          ByVal folderPath As String, _
                          ByVal baseName As String) As String
    lastOutcome = outcome
    lastMessages = Messenger.Messages()
    Messenger.Disarm

    On Error Resume Next
        WriteSummaryFile folderPath, baseName
    On Error GoTo 0

    CloseRun = outcome
End Function

'@sub-title Write the summary beside the file the run touched.
'@param folderPath String. The folder to write into.
'@param baseName String. The name the file is built on.
Private Sub WriteSummaryFile(ByVal folderPath As String, ByVal baseName As String)
    Dim filePath As String
    Dim fileNumber As Long

    If LenB(folderPath) = 0 Then Exit Sub
    If LenB(baseName) = 0 Then Exit Sub

    filePath = folderPath
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    filePath = filePath & baseName & SUMMARY_SUFFIX

    fileNumber = FreeFile

    On Error GoTo CloseFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, SetupLastSummary()
    Close #fileNumber
    Exit Sub

CloseFile:
    On Error Resume Next
        Close #fileNumber
    On Error GoTo 0
End Sub

'@sub-title Build the "ERROR <number>: <text>" line a failed run answers.
'@param errNumber Long. The error number, 0 for a refusal of the wrapper's own.
'@param message String. What went wrong.
'@return String. The outcome line.
Private Function RunError(ByVal errNumber As Long, ByVal message As String) As String
    RunError = OUTCOME_ERROR_LEAD & CStr(errNumber) & ": " & message
End Function

'@sub-title Whether a file sits at that path.
'@param filePath String. The path to look at.
'@return Boolean. True when the path names a file that is there.
Private Function FileIsThere(ByVal filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
        FileIsThere = (LenB(Dir$(filePath)) > 0)
    On Error GoTo 0
End Function

'@sub-title Whether a folder sits at that path.
'@details
'GetAttr rather than Dir$(path, vbDirectory). Dir$ with vbDirectory does not
'answer reliably for a folder on Mac Excel, and it cost this wrapper a red probe:
'a real folder was refused as missing and the export never started. GetAttr is
'what HeadlessBuild.IsFolder, TemporaryRepos and GenerationHost.FolderExists all
'use, so this is the idiom of the repo rather than a local choice.
'@param folderPath String. The path to look at.
'@return Boolean. True when the path names a folder that is there.
Private Function FolderIsThere(ByVal folderPath As String) As Boolean
    Dim folderAttributes As Long
    Dim trimmed As String

    trimmed = folderPath
    If LenB(trimmed) = 0 Then Exit Function
    If Right$(trimmed, 1) = Application.PathSeparator Then
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    End If

    On Error Resume Next
        folderAttributes = GetAttr(trimmed)
        If Err.Number = 0 Then
            FolderIsThere = ((folderAttributes And vbDirectory) = vbDirectory)
        End If
        Err.Clear
    On Error GoTo 0
End Function

'@sub-title The folder a path sits in.
'@param filePath String. A full file path.
'@return String. Everything before the last separator, empty when there is none.
Private Function ParentFolderOf(ByVal filePath As String) As String
    Dim cutAt As Long

    cutAt = InStrRev(filePath, Application.PathSeparator)
    If cutAt <= 1 Then Exit Function

    ParentFolderOf = Left$(filePath, cutAt - 1)
End Function

'@sub-title The name of a file with its folder and its extension taken off.
'@param filePath String. A file name or a full path.
'@return String. The bare name.
Private Function BaseNameOf(ByVal filePath As String) As String
    Dim bareName As String
    Dim cutAt As Long

    bareName = filePath

    cutAt = InStrRev(bareName, Application.PathSeparator)
    If cutAt > 0 Then bareName = Mid$(bareName, cutAt + 1)

    cutAt = InStrRev(bareName, ".")
    If cutAt > 1 Then bareName = Left$(bareName, cutAt - 1)

    BaseNameOf = bareName
End Function

'@sub-title What the export's summary file is named after.
'@details
'The export it wrote, and this workbook when it wrote nothing, so a refused
'export still leaves a readable file in the folder it was pointed at.
'@param exportPath String. What LastExportFile answered.
'@return String. The bare name to build the summary file on.
Private Function SummaryBase(ByVal exportPath As String) As String
    If LenB(exportPath) = 0 Then
        SummaryBase = BaseNameOf(ThisWorkbook.Name)
        Exit Function
    End If

    SummaryBase = BaseNameOf(exportPath)
End Function
