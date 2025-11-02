Attribute VB_Name = "SetupHelpers"

Option Explicit

Private Const PASSSHEETNAME As String = "__pass"
Private Const TRADSHEETNAME As String = "Translations"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const DICTSHEETNAME As String = "Dictionary"
Private Const CHOICESSHEETNAME As String = "Choices"
Private Const DROPDOWNSHEETNAME As String = "__variables"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const EXPORTSHEETNAME As String = "Exports"
Private Const TRANSLATIONSHEETNAME As String = "Translations"
Private Const CHECKINGSHEETNAME As String = "__checkRep"
Private Const ANALYSIS_TABLE_GLOBAL_SUMMARY As String = "Tab_global_summary"
Private Const ANALYSIS_TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const ANALYSIS_TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const ANALYSIS_TABLE_TS_DATA As String = "Tab_TimeSeries_Analysis"
Private Const ANALYSIS_TABLE_TS_GRAPH As String = "Tab_Graph_TimeSeries"
Private Const ANALYSIS_TABLE_TS_LABELS As String = "Tab_Label_TSGraph"
Private Const ANALYSIS_TABLE_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const ANALYSIS_TABLE_SPATIOTEMP As String = "Tab_SpatioTemporal_Analysis"
Private Const ANALYSIS_TABLE_SPATIOTEMP_SPECS As String = "Tab_SpatioTemporal_Specs"


'Start Rows and columns for dictionary, choices, and exports.
Private Const START_ROW_DICTIONARY As Long = 5
Private Const START_ROW_CHOICES As Long = 4
Private Const START_ROW_EXPORTS As Long = 4
Private Const START_COLUMN_DICTIONARY As Long = 1
Private Const START_COLUMN_CHOICES As Long = 1
Private Const START_COLUMN_EXPORTS As Long = 1

'Implement the password protection for the workbook entirely

'@section Basic Rows management in tables
'===============================================================================

'@sub-title Add or remove rows to a table
Public Sub ManageRows(ByVal sheetName As String, _
                      Optional ByVal del As Boolean = False, _
                      Optional ByVal allAnalysis As Boolean = False)
    Dim part As Object
    Dim targetSheet As Worksheet
    Dim dictSheet As Worksheet
    Dim dict As ILLdictionary
    Dim app As IApplicationState

    Set app = ApplicationState.Create(Application)
    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If (targetSheet Is Nothing) Then Exit Sub

    On Error GoTo Handler

    '5 is the start line of the dictionary
    '4 is the start column of the dictionary
    Select Case LCase$(Trim$(sheetName))
        Case "dictionary"
            Set part = ResolveDictionary(targetSheet)
        Case "choices"
            Set part = ResolveChoices(targetSheet)
        Case "analysis"
            If allAnalysis Then
                On Error Resume Next
                    targetSheet.Range("RNG_SelectTable").Value = "Add or remove rows of all tables"
                On Error GoTo 0
            End If
            Set part = ResolveAnalysis(targetSheet)
        Case "exports"
            Set dictSheet = ResolveSetupSheet("dict")
            Set part = LLExport.Create(targetSheet, START_ROW_EXPORTS, START_COLUMN_EXPORTS)
            Set dict = ResolveDictionary(dictSheet)
        Case Else
            Exit Sub
    End Select

    If Not (part Is Nothing) Then
        EnsureRowManagement sheetName, del, part, dict
    End If

    If Not (app Is Nothing) Then app.Restore
    Exit Sub
    
Handler:
    On Error Resume Next
    If Not (app Is Nothing) Then app.Restore
    ProtectSetupSheet sheetName
    If Err.Number <> 0 Then Debug.Print "Manage rows exited with an error: "; Err.Description; Err.Number 
End Sub


'@sub-title Ensure Row Management
Private Sub EnsureRowManagement(ByVal sheetName As String, ByVal del As Boolean, _ 
                                ByVal part As Object, Optional ByVal dict As ILLdictionary)


    UnProtectSetupSheet sheetName

    If dict Is Nothing Then
        part.ManageRows del
    Else
        UnProtectSetupSheet ResolveSetupSheetName("dict")
        part.ManageRows del, dict
        ProtectSetupSheet ResolveSetupSheetName("dict")
    End If
    ProtectSetupSheet sheetName
End Sub

'@sub-title Insert a list row at the active cell position
Public Sub InsertListRowAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim position As Long

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

    UnProtectSetupSheet sheetName

    If (lo.DataBodyRange Is Nothing) Then
        lo.ListRows.Add
    ElseIf (sheetName = ResolveSetupSheetName("ana")) Then
        targetSheet.Rows(targetCell.Row).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else 
        position = targetCell.Row - lo.HeaderRowRange.Row
        If position < 1 Or position > lo.ListRows.Count Then
            lo.ListRows.Add
        Else
            'InsertRow
            lo.ListRows.Add Position:=position           
        End If
    End If

    ProtectSetupSheet sheetName
End Sub

'@sub-title Delete the list row intersecting the active cell
Public Sub DeleteListRowAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim position As Long

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
    If lo.DataBodyRange Is Nothing Then Exit Sub

    position = targetCell.Row - lo.HeaderRowRange.Row
    If position < 1 Or position > lo.ListRows.Count Then Exit Sub

    UnProtectSetupSheet sheetName
        If sheetName <> ResolveSetupSheetName("ana") Then
            lo.ListRows(position).Delete
        Else
            targetSheet.Rows(targetCell.Row).Delete
        End If
    ProtectSetupSheet sheetName
End Sub

'@sub-title Delete the list column intersecting the active cell
Public Sub DeleteListColumnAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim colIndex As Long

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
    If colIndex < 1 Or colIndex > lo.ListColumns.Count Then Exit Sub

    UnProtectSetupSheet sheetName
        lo.ListColumns(colIndex).Delete
    ProtectSetupSheet sheetName
End Sub

'@section Filtering and Sorting tables
'===============================================================================

'@sub-title Clear filters on every listobject in the sheet
Public Sub ClearSheetFilters(ByVal sheetName As String)
    Dim targetSheet As Worksheet
    Dim lo As ListObject

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub

    UnProtectSetupSheet sheetName

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

    ProtectSetupSheet sheetName
End Sub

'@sub-title Sort setup tables based on the active worksheet
Public Sub SortSetupTables(ByVal sheetName As String)
    Dim targetSheet As Worksheet
    Dim normalizedName As String
    Dim choices As ILLChoices
    Dim ana As IAnalysis
    Dim lo As ListObject
    Dim tabl As ICustomTable
    Dim dict As ILLdictionary

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub

    normalizedName = LCase$(Trim$(sheetName))

    Select Case normalizedName
        Case "choices"
            Set choices = LLChoices.Create(targetSheet, START_ROW_CHOICES, START_COLUMN_CHOICES)
            UnProtectSetupSheet sheetName
                choices.Sort
            ProtectSetupSheet sheetName
        Case "analysis"
            Set ana = Analysis.Create(targetSheet)
            UnProtectSetupSheet sheetName
                ana.Sort
            ProtectSetupSheet sheetName
        Case "exports"
            On Error Resume Next
                Set lo = targetSheet.ListObjects(1)
            On Error GoTo 0
            If lo Is Nothing Then Exit Sub
            Set tabl = CustomTable.Create(lo)
            UnProtectSetupSheet sheetName
                tabl.Sort colName:="export number"
            ProtectSetupSheet sheetName
        Case "dictionary"
            Set dict = LLdictionary.Create(targetSheet, START_ROW_DICTIONARY, START_COLUMN_DICTIONARY)
            UnProtectSetupSheet sheetName
                dict.Sort

            ProtectSetupSheet sheetName
    End Select
End Sub

'@section Protect / UnProtect
'===============================================================================

'@sub-title Unprotect a worksheet
Public Sub UnProtectSetupSheet(ByVal sheetName As String)
    Dim pass As IPasswords
    Set pass = ResolveSetupPasswords()
    pass.UnProtect sheetName
End Sub

'@sub-title Protect a worksheet
Public Sub ProtectSetupSheet(ByVal sheetName As String)
    Dim pass As IPasswords
    Dim delRow As Boolean

    delRow = Not ((sheetName = TRADSHEETNAME) Or (sheetName = ANALYSISSHEETNAME))

    Set pass = ResolveSetupPasswords()
    pass.Protect sheetName, allowDeletingRows:=delRow
End Sub

'@section Translations
'===============================================================================

Public Sub ApplySetupTranslation(ByVal translator As ITranslationObject)
    Dim dictSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim analysisSheet As Worksheet
    Dim exportsSheet As Worksheet
    Dim dictionary As ILLdictionary
    Dim choices As ILLChoices
    Dim analysis As IAnalysis
    Dim exports As ILLExport
    Dim unlockDict As Boolean
    Dim unlockChoices As Boolean
    Dim unlockAnalysis As Boolean
    Dim unlockExports As Boolean

    On Error GoTo Cleanup

    Set dictSheet = ResolveSetupSheet("dict")
    If Not dictSheet Is Nothing Then
        UnProtectSetupSheet DICTSHEETNAME
        unlockDict = True
        Set dictionary = ResolveDictionary(dictSheet)
        dictionary.Translate translator
        ProtectSetupSheet DICTSHEETNAME
        unlockDict = False
    End If

    Set choicesSheet = ResolveSetupSheet("choi")
    If Not choicesSheet Is Nothing Then
        UnProtectSetupSheet CHOICESSHEETNAME
        unlockChoices = True
        Set choices = ResolveChoices(choicesSheet)
        choices.Translate translator
        ProtectSetupSheet CHOICESSHEETNAME
        unlockChoices = False
    End If

    Set analysisSheet = ResolveSetupSheet("ana")
    If Not analysisSheet Is Nothing Then
        UnProtectSetupSheet ANALYSISSHEETNAME
        unlockAnalysis = True
        Set analysis = ResolveAnalysis(analysisSheet)
        analysis.Translate translator
        ProtectSetupSheet ANALYSISSHEETNAME
        unlockAnalysis = False
    End If

    Set exportsSheet = ResolveSetupSheet("exp")
    If Not exportsSheet Is Nothing Then
        UnProtectSetupSheet EXPORTSHEETNAME
        unlockExports = True
        Set exports = LLExport.Create(exportsSheet, START_ROW_EXPORTS, START_COLUMN_EXPORTS)
        exports.Translate translator
        ProtectSetupSheet EXPORTSHEETNAME
        unlockExports = False
    End If

Cleanup:
    If unlockDict Then ProtectSetupSheet DICTSHEETNAME
    If unlockChoices Then ProtectSetupSheet CHOICESSHEETNAME
    If unlockAnalysis Then ProtectSetupSheet ANALYSISSHEETNAME
    If unlockExports Then ProtectSetupSheet EXPORTSHEETNAME
    If Err.Number <> 0 Then Err.Raise Err.Number, "SetupHelpers.ApplySetupTranslation", Err.Description
End Sub


Public Function ResolveSetupSheetName(ByVal sheetKey As String) As String
    Dim normalized As String

    normalized = LCase$(Trim$(sheetKey))

    Select Case normalized
        Case "dict"
            ResolveSetupSheetName = DICTSHEETNAME
        Case "choi"
            ResolveSetupSheetName = CHOICESSHEETNAME
        Case "ana"
            ResolveSetupSheetName = ANALYSISSHEETNAME
        Case "trans"
            ResolveSetupSheetName = TRANSLATIONSHEETNAME
        Case "exp"
            ResolveSetupSheetName = EXPORTSHEETNAME
        Case "drop"
            ResolveSetupSheetName = DROPDOWNSHEETNAME
        Case "check"
            ResolveSetupSheetName = CHECKINGSHEETNAME
    End Select
End Function

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

'Prepare the Import Form
Public Sub PrepareImportsForm(Optional ByVal cleanSetup As Boolean = False)
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

'Import the setup from 

Public Sub ImportOrCleanSetup()
    Const CLEAN_CAPTION As String = "Clear"
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
    Dim importCaption As String
    Dim isClean As Boolean
    Dim importPath As String
    Dim servicePath As String
    Dim service As ISetupImportService
    Dim pass As IPasswords
    Dim app As IApplicationState
    Dim sheets As BetterArray
    Dim infoText As String
    Dim completed As Boolean

    On Error GoTo Handler

    Set formRef = [Imports]
    If formRef Is Nothing Then Exit Sub

    importDict = CBool(formRef.DictionaryCheck.Value)
    importChoi = CBool(formRef.ChoiceCheck.Value)
    importExp = CBool(formRef.ExportsCheck.Value)
    importAna = CBool(formRef.AnalysisCheck.Value)
    importTrans = CBool(formRef.TranslationsCheck.Value)
    conformityCheck = CBool(formRef.ConformityCheck.Value)
    Set progressLabel = formRef.LabProgress
    importCaption = Trim$(CStr(formRef.DoButton.Caption))
    isClean = (StrComp(importCaption, CLEAN_CAPTION, vbTextCompare) = 0)

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
    Set app = ApplicationState.Create(Application)

    app.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set service = SetupImportService.Create(servicePath, progressLabel)
    service.Check importDict, importChoi, importExp, importAna, importTrans, cleanSetup:=isClean

    If isClean Then
        infoText = ExecuteCleanOperation(service, pass, sheets, CLEAN_DONE, ABORTED)
    Else
        infoText = ExecuteImportOperation(service, pass, sheets, conformityCheck, IMPORT_DONE)
    End If
    completed = True

Cleanup:
    If Not app Is Nothing Then app.Restore

    If completed Then 
        formRef.Hide
        If conformityCheck And Not isClean Then
            On Error Resume Next
                ThisWorkbook.Worksheets(CHECKINGSHEETNAME).Activate
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
Private Function BuildSelectedSheets(ByVal importDict As Boolean, _
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
    If importTrans Then sheets.Push TRANSLATIONSHEETNAME

    Set BuildSelectedSheets = sheets
End Function

'@Description("Extract the import path from the form label caption")
Private Function ParseImportPath(ByVal captionText As String) As String
    Dim trimmed As String

    trimmed = Replace(captionText, "Path:", vbNullString, 1, 1, vbTextCompare)
    ParseImportPath = Trim$(trimmed)
End Function

'@Description("Execute the workbook-driven import using the selected sheets")
Private Function ExecuteImportOperation(ByVal service As ISetupImportService, _
                                        ByVal pass As IPasswords, _
                                        ByVal sheets As BetterArray, _
                                        ByVal runConformityCheck As Boolean, _
                                        ByVal successMessage As String) As String
    
    
    service.ImportFromWorkbook pass, sheets
    If runConformityCheck Then CheckTheSetup

    PostMaintenanceAfterImport

    ExecuteImportOperation = successMessage
End Function

'@Description("Execute the clean workflow against selected sheets")
Private Function ExecuteCleanOperation(ByVal service As ISetupImportService, _
                                       ByVal pass As IPasswords, _
                                       ByVal sheets As BetterArray, _
                                       ByVal successMessage As String, _
                                       ByVal abortMessage As String) As String
    Const CLEAR_PROMPT As String = "Do you really want to clear the setup?"

    Dim confirmation As VbMsgBoxResult
    Dim idx As Long
    Dim sheetName As String

    confirmation = MsgBox(CLEAR_PROMPT, vbYesNo + vbQuestion, "Confirmation")
    If confirmation <> vbYes Then
        ExecuteCleanOperation = abortMessage
        Exit Function
    End If

    service.Clean pass, sheets

    For idx = sheets.LowerBound To sheets.UpperBound
        sheetName = CStr(sheets.Item(idx))
        If StrComp(sheetName, ANALYSISSHEETNAME, vbTextCompare) = 0 Then
            ManageRows sheetName, del:=True, allAnalysis:=True
        Else
            ManageRows sheetName, del:=True
        End If
    Next idx

    On Error Resume Next
        ThisWorkbook.Worksheets("__checkRep").Cells.Clear
    On Error GoTo 0

    ExecuteCleanOperation = successMessage
End Function

Public Sub PostMaintenanceAfterImport()
    Dim prep As ISetupPreparation
    
    Set prep = SetupPreparation.Create(ThisWorkbook)
    prep.EnsureUpdatedRegistry

    SetupEventsManager.ResetTranslationCounter
    SetupEventsManager.RefreshAnalysisDropdowns forceUpdate:=True
    SetupEventsManager.RecalculateAnalysis
End Sub

'@section Checkings
'===============================================================================

'Factory helpers
'-------------------------------------------------------------------------------
'@sub-title Resolve the workbook that will be analysed.
'@param hostBook Optional workbook reference. Defaults to ThisWorkbook.
'@return Workbook reference guaranteed to be non-Nothing.
Private Function ResolveWorkbook(Optional ByVal hostBook As Workbook) As Workbook
    If hostBook Is Nothing Then
        Set hostBook = ThisWorkbook
    End If

    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required"
    End If

    Set ResolveWorkbook = hostBook
End Function

'@sub-title Instantiate a SetupErrors checker.
'@param hostBook Workbook containing setup sheets to inspect.
'@return ISetupErrors instance ready to execute.
Private Function CreateChecker(Optional ByVal hostBook As Workbook) As ISetupErrors
    Dim checker As ISetupErrors

    Set checker = SetupErrors.Create(ResolveWorkbook(hostBook))
    Set CreateChecker = checker
End Function

'@section Public API
'-------------------------------------------------------------------------------
'@sub-title Backwards-compatible entry point matching the legacy module signature.
Public Sub CheckTheSetup()
    RunSetupChecks ThisWorkbook
End Sub

'@sub-title Execute setup checks against the provided workbook.
'@param hostBook Optional workbook. When omitted, ThisWorkbook is used.
Public Sub RunSetupChecks(Optional ByVal hostBook As Workbook)
    Dim checker As ISetupErrors
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo RunFailed
        Set checker = CreateChecker(hostBook)
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

'@Description("Build the default sheet list consumed by table-driven imports")
Public Function DefaultSetupSheets() As BetterArray
    Dim sheets As BetterArray

    Set sheets = New BetterArray
    sheets.LowerBound = 1
    sheets.Push DICTSHEETNAME, CHOICESSHEETNAME, EXPORTSHEETNAME, _
                 ANALYSISSHEETNAME, TRANSLATIONSHEETNAME

    Set DefaultSetupSheets = sheets
End Function

'@Description("Prompt user to pick an import workbook and return its path")
Public Function SelectSetupImportPath(ByVal filters As String) As String
    Dim io As IOSFiles

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
Public Function ResolveSetupPasswords() As IPasswords
    Dim passSheet As Worksheet
    Set passSheet = ThisWorkbook.Worksheets(PASSSHEETNAME)
    Set ResolveSetupPasswords = Passwords.Create(passSheet)
End Function

Public Function ResolveUpdatedValues() As IUpdatedValues
    Set ResolveUpdatedValues = UpdatedValues.Create(ResolveRegistrySheet())
End Function

Public Function ResolveDictionary(Optional ByVal hostSheet As Worksheet) As ILLdictionary
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveSetupSheet("dict")
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveDictionary = LLdictionary.Create(targetSheet, START_ROW_DICTIONARY, START_COLUMN_DICTIONARY)
End Function

Public Function ResolveChoices(Optional ByVal hostSheet As Worksheet) As ILLChoices
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveSetupSheet("choi")
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveChoices = LLChoices.Create(targetSheet, START_ROW_CHOICES, START_COLUMN_CHOICES)
End Function

Public Function ResolveAnalysis(Optional ByVal hostSheet As Worksheet) As IAnalysis
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveSetupSheet("ana")
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveAnalysis = Analysis.Create(targetSheet)
End Function

Public Function ResolveVariables(Optional ByVal dictionary As ILLdictionary, _
                                 Optional ByVal hostSheet As Worksheet) As ILLVariables
    Dim dict As ILLdictionary

    If dictionary Is Nothing Then
        Set dict = ResolveDictionary(hostSheet)
    Else
        Set dict = dictionary
    End If

    If dict Is Nothing Then Exit Function

    Set ResolveVariables = LLVariables.Create(dict)
End Function

Public Function ResolveDropdowns(Optional ByVal hostSheet As Worksheet, _
                                 Optional ByVal headerPrefix As String = "dropdown_") As IDropdownLists
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveSetupSheet("drop")
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveDropdowns = DropdownLists.Create(targetSheet, headerPrefix)
End Function

Public Function ResolveAnalysisTableName(ByVal tableKey As String) As String
    Dim normalized As String

    normalized = LCase$(Trim$(tableKey))

    Select Case normalized
        Case "global", "global_summary", "summary", "analysis_summary"
            ResolveAnalysisTableName = ANALYSIS_TABLE_GLOBAL_SUMMARY
        Case "univariate", "uni", "univariate_analysis"
            ResolveAnalysisTableName = ANALYSIS_TABLE_UNIVARIATE
        Case "bivariate", "bi", "bivariate_analysis"
            ResolveAnalysisTableName = ANALYSIS_TABLE_BIVARIATE
        Case "ts", "timeseries", "time_series", "time-series", "ts_table", _
             "ts_data", "analysis_ts", "timeseries_analysis"
            ResolveAnalysisTableName = ANALYSIS_TABLE_TS_DATA
        Case "ts_graph", "graph_ts", "graph", "chart", "time_series_graph", _
             "timeseries_graph", "analysis_graph"
            ResolveAnalysisTableName = ANALYSIS_TABLE_TS_GRAPH
        Case "ts_labels", "graph_labels", "labels", "ts_graph_labels", _
             "graph_titles", "ts_titles", "analysis_graph_labels"
            ResolveAnalysisTableName = ANALYSIS_TABLE_TS_LABELS
        Case "spatial", "spatial_analysis", "geo_spatial", "geospatial", "spatial_table"
            ResolveAnalysisTableName = ANALYSIS_TABLE_SPATIAL
        Case "spatiotemporal", "spatio-temporal", "spatiotemporal_analysis", _
             "spatio", "spatiotemp", "spatio_temp", "st"
            ResolveAnalysisTableName = ANALYSIS_TABLE_SPATIOTEMP
        Case "spatiotemporal_specs", "spatio-temporal_specs", "spatiotemp_specs", _
             "spatial_specs", "spatial_specifications", "analysis_specs", "spatial_tables_specs"
            ResolveAnalysisTableName = ANALYSIS_TABLE_SPATIOTEMP_SPECS
        Case Else
            ResolveAnalysisTableName = tableKey
    End Select
End Function

Public Function ResolveAnalysisTable(ByVal tableKey As String, _
                                     Optional ByVal hostSheet As Worksheet, _
                                     Optional ByVal idColumn As String = vbNullString, _
                                     Optional ByVal idPrefix As String = vbNullString) As ICustomTable
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim tableName As String

    If LenB(tableKey) = 0 Then Exit Function

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveSetupSheet("ana")
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    tableName = ResolveAnalysisTableName(tableKey)

    On Error Resume Next
        Set lo = targetSheet.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    Set ResolveAnalysisTable = CustomTable.Create(lo, idColumn, idPrefix)
End Function
