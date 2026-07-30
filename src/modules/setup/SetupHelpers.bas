Attribute VB_Name = "SetupHelpers"

Option Explicit

'@Folder("Setup")
'@ModuleDescription("The import and clean flow of the setup, and the accessors its form and ribbon share")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString

'This module owns the import and clean flow and the few accessors the Imports
'form and the ribbon share. Row management, sorting, sheet protection and the
'setup translation all live on EventSetup now, and the ribbon reaches them
'through SetupEventsManager.EventSetupService.
'
'THE SHEET NAMES BELOW ARE DECLARED TWICE ON PURPOSE
'-------------------------------------------------------------------------------
'EventSetup.cls declares the same names for the event and row work. Each file
'keeps the constants it needs, so neither depends on the other for a string.
'Change one and change the other.

Private Const PASSSHEETNAME As String = "__pass"
Private Const TRADSHEETNAME As String = "Translations"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const DICTSHEETNAME As String = "Dictionary"
Private Const CHOICESSHEETNAME As String = "Choices"
Private Const DROPDOWNSHEETNAME As String = "__variables"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const EXPORTSHEETNAME As String = "Exports"
Private Const CHECKINGSHEETNAME As String = "__checkRep"

'Cached password helper (lazily created once per VBA session)
Private cachedPasswords As Passwords

'Which job the Imports form was prepared for. PrepareImportsForm writes it and
'ImportOrCleanSetup reads it. The button caption used to carry this, so
'translating the form or editing the caption changed what the button did.
Private cleanModeSelected As Boolean

'@section Basic Rows management in tables
'===============================================================================

'@sub-title Delete the list column intersecting the active cell
'@details
'The caller confirms with the user and checks the sheet before calling this. Only
'the Translations sheet has columns a user may remove, and the ribbon hides the
'button everywhere else.
'@param sheetName String. Sheet holding the table.
'@param targetCell Range. Cell inside the column to remove.
Public Sub DeleteListColumnAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim colIndex As Long
    Dim svc As EventSetup

    If targetCell Is Nothing Then Exit Sub

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If Not targetCell.Parent Is targetSheet Then Exit Sub

    On Error Resume Next
        Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    colIndex = targetCell.Column - lo.Range.Column + 1
    If (colIndex <= 1) Or colIndex > lo.ListColumns.Count Then Exit Sub

    Set svc = SetupEventsManager.EventSetupService

    svc.UnprotectSetupSheet sheetName
        lo.ListColumns(colIndex).Delete
    svc.ProtectSetupSheet sheetName
End Sub

'@section Sheet name resolution
'===============================================================================

'@sub-title Turn a short sheet key into the sheet name the workbook carries
'@param sheetKey String. Short key or full sheet name.
'@return String. The sheet name, or empty when the key is unknown.
Public Function ResolveSetupSheetName(ByVal sheetKey As String) As String
    Dim normalized As String

    normalized = LCase$(Trim$(sheetKey))

    Select Case normalized
        Case "dict", "dictionary"
            ResolveSetupSheetName = DICTSHEETNAME
        Case "choi", "choice", "choices"
            ResolveSetupSheetName = CHOICESSHEETNAME
        Case "ana", "analysis"
            ResolveSetupSheetName = ANALYSISSHEETNAME
        Case "trans", "translation", "translations"
            ResolveSetupSheetName = TRADSHEETNAME
        Case "exp", "exports", "export"
            ResolveSetupSheetName = EXPORTSHEETNAME
        Case "drop", "dropdowns", "dropdown"
            ResolveSetupSheetName = DROPDOWNSHEETNAME
        Case "check", "checking", "checkings"
            ResolveSetupSheetName = CHECKINGSHEETNAME
    End Select
End Function

'@sub-title Resolve a sheet from a short key or a full name
'@param sheetKey String. Short key or full sheet name.
'@return Worksheet. The worksheet, or Nothing when it is absent.
Public Function ResolveSetupSheet(ByVal sheetKey As String) As Worksheet
    Dim resolvedName As String

    resolvedName = ResolveSetupSheetName(sheetKey)
    If LenB(resolvedName) = 0 Then resolvedName = sheetKey

    On Error Resume Next
        Set ResolveSetupSheet = ThisWorkbook.Worksheets(resolvedName)
    On Error GoTo 0
End Function

'@section Imports/Exports
'===============================================================================

'@sub-title Lay the Imports form out for the job it is about to do
'@param cleanSetup Optional Boolean. True prepares the clear job, False the import job.
Public Sub PrepareImportsForm(Optional ByVal cleanSetup As Boolean = False)
    cleanModeSelected = cleanSetup

    If cleanSetup Then
        [Imports].LoadButton.Visible = False
        [Imports].LabPath.Visible = False
        [Imports].InfoChoice.Caption = "Select what to Clear"
        [Imports].DictionaryCheck.Caption = "Clear Dictionary"
        [Imports].ChoiceCheck.Caption = "Clear Choices"
        [Imports].ExportsCheck.Caption = "Clear Exports"
        [Imports].AnalysisCheck.Caption = "Clear Analysis"
        [Imports].TranslationsCheck.Caption = "Clear Translation"
        [Imports].ConformityCheck.Visible = False
        [Imports].DoButton.Caption = "Clear"

        'Resize and change position of elements
        [Imports].Height = 400
        [Imports].InfoChoice.Top = 20
        [Imports].DictionaryCheck.Top = 50
        [Imports].ChoiceCheck.Top = 80
        [Imports].ExportsCheck.Top = 110
        [Imports].AnalysisCheck.Top = 140
        [Imports].TranslationsCheck.Top = 170
        [Imports].LabProgress.Top = 200
        [Imports].DoButton.Top = 270
        [Imports].Quit.Top = 310
    Else
        [Imports].InfoChoice.Caption = "Select what to Import"
        [Imports].DictionaryCheck.Caption = "Import Dictionary"
        [Imports].ChoiceCheck.Caption = "Import Choices"
        [Imports].ExportsCheck.Caption = "Import Exports"
        [Imports].AnalysisCheck.Caption = "Import Analysis"
        [Imports].TranslationsCheck.Caption = "Import Translation"
        [Imports].ConformityCheck.Visible = True
        [Imports].LoadButton.Visible = True
        [Imports].LabPath.Visible = True
        [Imports].DoButton.Caption = "Import"

        'resize the worksheet and position of elements
        [Imports].Height = 500
        [Imports].LoadButton.Top = 10
        [Imports].LabPath.Top = 55
        [Imports].InfoChoice.Top = 135
        [Imports].DictionaryCheck.Top = 170
        [Imports].ChoiceCheck.Top = 200
        [Imports].ExportsCheck.Top = 230
        [Imports].AnalysisCheck.Top = 260
        [Imports].TranslationsCheck.Top = 290
        [Imports].DoButton.Top = 350
        [Imports].LabProgress.Top = 390
        [Imports].Quit.Top = 440
    End If
End Sub

'@sub-title Run the import or the clean the form was prepared for
Public Sub ImportOrCleanSetup()
    Const IMPORT_DONE As String = "Import Done!"
    Const CLEAN_DONE As String = "Setup cleared!"
    Const ABORTED As String = "Aborted!"

    Dim formRef As Imports
    Dim importDict As Boolean
    Dim importChoi As Boolean
    Dim importExp As Boolean
    Dim importAna As Boolean
    Dim importTrans As Boolean
    Dim conformityCheck As Boolean
    Dim progressLabel As Object
    Dim isClean As Boolean
    Dim importPath As String
    Dim servicePath As String
    Dim service As SetupImport
    Dim pass As Passwords
    Dim sheets As BetterArray
    Dim infoText As String
    Dim completed As Boolean
    Dim originalSheet As Worksheet

    On Error GoTo Handler

    Set originalSheet = ActiveSheet
    Set formRef = [Imports]
    If formRef Is Nothing Then Exit Sub

    importDict = CBool(formRef.DictionaryCheck.Value)
    importChoi = CBool(formRef.ChoiceCheck.Value)
    importExp = CBool(formRef.ExportsCheck.Value)
    importAna = CBool(formRef.AnalysisCheck.Value)
    importTrans = CBool(formRef.TranslationsCheck.Value)
    conformityCheck = CBool(formRef.ConformityCheck.Value)
    Set progressLabel = formRef.LabProgress
    isClean = cleanModeSelected

    If isClean Then conformityCheck = False

    importPath = ParseImportPath(formRef.LabPath.Caption)
    infoText = ABORTED
    progressLabel.Caption = vbNullString

    If Not isClean And LenB(importPath) = 0 Then
        MsgBox "Select a setup workbook before importing.", vbExclamation
        Exit Sub
    End If

    If (Not isClean) Then
        servicePath = importPath
    Else
        servicePath = ThisWorkbook.FullName
    End If

    Set sheets = BuildSelectedSheets(importDict, importChoi, importExp, importAna, importTrans)
    Set pass = ResolveSetupPasswords()
    SetupEventsManager.EnterBusyState calculateOnSave:=False

    Set service = SetupImport.Create(servicePath, progressLabel)
    service.Check importDict, importChoi, importExp, importAna, importTrans, cleanSetup:=isClean

    If isClean Then
        infoText = ExecuteCleanOperation(service, pass, sheets, CLEAN_DONE, ABORTED)
    Else
        infoText = ExecuteImportOperation(service, pass, sheets, conformityCheck, IMPORT_DONE)
    End If
    completed = True

Cleanup:
    SetupEventsManager.ExitBusyState

    If completed Then
        formRef.Hide
        If conformityCheck And Not isClean Then
            On Error Resume Next
                ThisWorkbook.Worksheets(CHECKINGSHEETNAME).Activate
            On Error GoTo 0
        ElseIf Not originalSheet Is Nothing Then
            On Error Resume Next
                originalSheet.Activate
            On Error GoTo 0
        End If
        MsgBox infoText
    End If
    Exit Sub

Handler:
    Debug.Print "SetupHelpers.ImportOrCleanSetup: "; Err.Number; Err.Description
    MsgBox "Failed to process the setup import/clean: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'@Description("Build the sheet list based on selected options")
Public Function BuildSelectedSheets(ByVal importDict As Boolean, _
                                     ByVal importChoi As Boolean, _
                                     ByVal importExp As Boolean, _
                                     ByVal importAna As Boolean, _
                                     ByVal importTrans As Boolean) As BetterArray
    Dim sheets As BetterArray

    Set sheets = New BetterArray
    sheets.LowerBound = 1

    If importDict Then sheets.Push DICTSHEETNAME
    If importChoi Then sheets.Push CHOICESSHEETNAME
    If importExp Then sheets.Push EXPORTSHEETNAME
    If importAna Then sheets.Push ANALYSISSHEETNAME
    If importTrans Then sheets.Push TRADSHEETNAME

    Set BuildSelectedSheets = sheets
End Function

'@Description("Extract the import path from the form label caption")
Private Function ParseImportPath(ByVal captionText As String) As String
    Dim trimmed As String

    trimmed = Replace(captionText, "Path:", vbNullString, 1, 1, vbTextCompare)
    ParseImportPath = Trim$(trimmed)
End Function

'@Description("Execute the workbook-driven import using the selected sheets")
Private Function ExecuteImportOperation(ByVal service As SetupImport, _
                                        ByVal pass As Passwords, _
                                        ByVal sheets As BetterArray, _
                                        ByVal runConformityCheck As Boolean, _
                                        ByVal successMessage As String) As String


    service.Import pass, sheets
    If runConformityCheck Then CheckTheSetup

    PostImportMaintenance

    ExecuteImportOperation = successMessage
End Function

'@Description("Execute the clean workflow against selected sheets")
Private Function ExecuteCleanOperation(ByVal service As SetupImport, _
                                       ByVal pass As Passwords, _
                                       ByVal sheets As BetterArray, _
                                       ByVal successMessage As String, _
                                       ByVal abortMessage As String) As String
    Const CLEAR_PROMPT As String = "Do you really want to clear the setup?"

    Dim confirmation As VbMsgBoxResult
    Dim idx As Long
    Dim sheetName As String
    Dim svc As EventSetup

    confirmation = MsgBox(CLEAR_PROMPT, vbYesNo + vbQuestion, "Confirmation")
    If confirmation <> vbYes Then
        ExecuteCleanOperation = abortMessage
        Exit Function
    End If

    service.Clean pass, sheets

    'The clean emptied whole sheets, so the managers the service cached before it
    'ran were built against columns those sheets may no longer carry.
    SetupEventsManager.ResetEventSetupCaches
    Set svc = SetupEventsManager.EventSetupService

    For idx = sheets.LowerBound To sheets.UpperBound
        sheetName = CStr(sheets.Item(idx))
        If StrComp(sheetName, ANALYSISSHEETNAME, vbTextCompare) = 0 Then
            SelectAllAnalysisTables sheetName
        End If
        svc.ManageRows sheetName, del:=True
    Next idx

    On Error Resume Next
        ThisWorkbook.Worksheets(CHECKINGSHEETNAME).Cells.Clear
    On Error GoTo 0

    ExecuteCleanOperation = successMessage
End Function

'@sub-title Point the Analysis table selector at every table before a clean
'@details
'EventSetup.ManageRows reads RNG_SelectTable to learn which analysis table the
'user means. The clean means all of them.
'@param sheetName String. The Analysis sheet name.
Private Sub SelectAllAnalysisTables(ByVal sheetName As String)
    Dim targetSheet As Worksheet

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
        If Not targetSheet Is Nothing Then
            targetSheet.Range("RNG_SelectTable").Value = "Add or remove rows of all tables"
        End If
    On Error GoTo 0
End Sub

'@sub-title Rebuild the watcher registry and the analysis dropdowns after an import
'@details
'One busy pair covers the whole job. The three manager routines below each enter
'a state of their own, and busyDepth makes that nesting safe, so this is one
'restore instead of three.
Public Sub PostImportMaintenance()
    Dim prep As SetupPreparation
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Cleanup
    SetupEventsManager.EnterBusyState calculateOnSave:=False

    'The import rewrote whole sheets, so the managers the service cached before
    'it ran were built against columns those sheets may no longer carry.
    SetupEventsManager.ResetEventSetupCaches

    Set prep = SetupPreparation.Create(ThisWorkbook)
    prep.ResetUpdatedRegistry

    SetupEventsManager.ResetTranslationCounter
    SetupEventsManager.RefreshAnalysisDropdowns forceUpdate:=True
    SetupEventsManager.RecalculateAnalysis

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description
    SetupEventsManager.ExitBusyState
    If errNumber <> 0 Then
        Err.Raise errNumber, "SetupHelpers.PostImportMaintenance", errDescription
    End If
End Sub

'@section Checkings
'===============================================================================

'@sub-title Execute setup checks against the provided workbook.
'@param hostBook Optional workbook. When omitted, ThisWorkbook is used.
Public Sub CheckTheSetup(Optional ByVal hostBook As Workbook)

    Dim checker As SetupErrors
    Dim targetBook As Workbook
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set targetBook = hostBook
    If targetBook Is Nothing Then Set targetBook = ThisWorkbook

    On Error GoTo RunFailed
        Set checker = SetupErrors.Create(targetBook)
        checker.Run
    Exit Sub

RunFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
End Sub

'@section Helpers
'===============================================================================

'@Description("Prompt user to pick an import workbook and return its path")
Public Function SelectSetupImportPath(ByVal filters As String) As String
    Dim io As OSFiles

    Set io = OSFiles.Create()
    io.LoadFile filters

    If io.HasValidFile() Then
        SelectSetupImportPath = Trim$(CStr(io.File))
    End If
End Function

'@sub-title Retrieve the translations listobject when available
Public Function ResolveTranslationsTable() As ListObject
    Dim sh As Worksheet

    Set sh = ResolveSetupSheet("trans")
    If sh Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveTranslationsTable = sh.ListObjects(TABTRANSLATION)
    On Error GoTo 0
End Function

'@sub-title Retrieve the registry worksheet capturing updated values
Public Function ResolveRegistrySheet() As Worksheet
    On Error Resume Next
        Set ResolveRegistrySheet = ThisWorkbook.Worksheets(UPDATEDSHEETNAME)
    On Error GoTo 0
End Function


'@Description("Provide the password manager used for setup protections")
Public Function ResolveSetupPasswords() As Passwords
    If cachedPasswords Is Nothing Then
        Dim passSheet As Worksheet
        Set passSheet = ThisWorkbook.Worksheets(PASSSHEETNAME)
        Set cachedPasswords = Passwords.Create(passSheet)
    End If
    Set ResolveSetupPasswords = cachedPasswords
End Function

Public Function ResolveUpdatedValues() As UpdatedValues
    Set ResolveUpdatedValues = UpdatedValues.Create(ResolveRegistrySheet())
End Function
