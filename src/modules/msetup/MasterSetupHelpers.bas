Attribute VB_Name = "MasterSetupHelpers"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Utility helpers shared across master setup modules.")
'@depends DropdownLists, CustomTable, Passwords, BetterArray, MasterSetupVariables, LLdictionary, LLVariables, LLChoices, UpdatedValues, Development, ApplicationState, HiddenNames

Private Const VARIABLES_SHEETNAME As String = "Variables"
Private Const TRANSLATIONS_SHEETNAME As String = "Translations"
Private Const CHOICES_SHEETNAME As String = "Choices"
Private Const DROPDOWNS_SHEETNAME As String = "__dropdowns"
Private Const REGISTRY_SHEETNAME As String = "__updated"
Private Const PASSWORDS_SHEETNAME As String = "__pass"
Private Const DEVELOPMENT_SHEETNAME As String = "Dev"
Private Const CONFIG_SHEETS_LIST As String = "__configSheets"

Private Const START_ROW_VARIABLES As Long = 5
Private Const START_COLUMN_VARIABLES As Long = 1
Private Const START_ROW_CHOICES As Long = 4
Private Const START_COLUMN_CHOICES As Long = 1
Private Const DEFAULT_DROPDOWN_PREFIX As String = "dropdown_"
Private Const SHEET_TAG_NAME As String = "sheetTag"
Private Const DISEASE_TAG_VALUE As String = "disease"
Private Const VARIABLES_TAG_VALUE As String = "variables"
Private Const CHOICES_TAG_VALUE As String = "choices"
Private Const DEFAULT_ROW_BATCH As Long = 5
'A threshold of 0 lets CustomTable count the formula columns of the first row
'itself, so a line carrying only its formulas counts as empty on every sheet.
Private Const EMPTY_ROW_THRESHOLD As Long = 0

'@section Workbook helpers
'===============================================================================
Public Function ResolveMasterSetupWorkbook(Optional ByVal hostBook As Workbook) As Workbook
    Set ResolveMasterSetupWorkbook = EnsureWorkbook(hostBook)
End Function

Private Function EnsureWorkbook(Optional ByVal hostBook As Workbook) As Workbook
    If hostBook Is Nothing Then
        Set hostBook = ThisWorkbook
    End If

    If hostBook Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, "MasterSetupHelpers.EnsureWorkbook"
    End If

    Set EnsureWorkbook = hostBook
End Function

'@section Worksheet lookup
'===============================================================================
Public Function ResolveMasterSetupSheetName(ByVal sheetKey As String) As String
    Select Case LCase$(Trim$(sheetKey))
        Case "vars", "variables"
            ResolveMasterSetupSheetName = VARIABLES_SHEETNAME
        Case "trans", "translations"
            ResolveMasterSetupSheetName = TRANSLATIONS_SHEETNAME
        Case "choi", "choices"
            ResolveMasterSetupSheetName = CHOICES_SHEETNAME
        Case "drop", "dropdowns"
            ResolveMasterSetupSheetName = DROPDOWNS_SHEETNAME
        Case "reg", "registry"
            ResolveMasterSetupSheetName = REGISTRY_SHEETNAME
        Case "pass", "passwords"
            ResolveMasterSetupSheetName = PASSWORDS_SHEETNAME
        Case "dev", "development"
            ResolveMasterSetupSheetName = DEVELOPMENT_SHEETNAME
        Case Else
            ResolveMasterSetupSheetName = sheetKey
    End Select
End Function

Public Function ResolveMasterSetupSheet(ByVal sheetKey As String, _
                                        Optional ByVal hostBook As Workbook) As Worksheet
    Dim resolvedName As String
    Dim targetWorkbook As Workbook

    resolvedName = ResolveMasterSetupSheetName(sheetKey)
    Set targetWorkbook = ResolveMasterSetupWorkbook(hostBook)

    On Error Resume Next
        Set ResolveMasterSetupSheet = targetWorkbook.Worksheets(resolvedName)
    On Error GoTo 0
End Function

Public Function ResolveMasterVariablesSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterVariablesSheet = ResolveMasterSetupSheet("vars", hostBook)
End Function

Public Function ResolveMasterTranslationsSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterTranslationsSheet = ResolveMasterSetupSheet("trans", hostBook)
End Function

Public Function ResolveMasterChoicesSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterChoicesSheet = ResolveMasterSetupSheet("choi", hostBook)
End Function

Public Function ResolveMasterDropdownsSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterDropdownsSheet = ResolveMasterSetupSheet("drop", hostBook)
End Function

Public Function ResolveMasterRegistrySheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterRegistrySheet = ResolveMasterSetupSheet("reg", hostBook)
End Function

Public Function ResolveMasterPasswordsSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterPasswordsSheet = ResolveMasterSetupSheet("pass", hostBook)
End Function

Public Function ResolveMasterDevelopmentSheet(Optional ByVal hostBook As Workbook) As Worksheet
    Set ResolveMasterDevelopmentSheet = ResolveMasterSetupSheet("dev", hostBook)
End Function

'@section Class factories
'===============================================================================
Public Function ResolveMasterDictionary(Optional ByVal hostSheet As Worksheet) As LLdictionary
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterVariablesSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterDictionary = LLdictionary.Create(targetSheet, START_ROW_VARIABLES, START_COLUMN_VARIABLES)
End Function

Public Function ResolveMasterChoices(Optional ByVal hostSheet As Worksheet) As LLChoices
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterChoicesSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterChoices = LLChoices.Create(targetSheet, START_ROW_CHOICES, START_COLUMN_CHOICES)
End Function

Public Function ResolveMasterVariables(Optional ByVal dictionary As LLdictionary, _
                                       Optional ByVal hostSheet As Worksheet) As LLVariables
    Dim resolvedDictionary As LLdictionary

    If dictionary Is Nothing Then
        If hostSheet Is Nothing Then
            Set resolvedDictionary = ResolveMasterDictionary()
        Else
            Set resolvedDictionary = ResolveMasterDictionary(hostSheet)
        End If
    Else
        Set resolvedDictionary = dictionary
    End If

    If resolvedDictionary Is Nothing Then Exit Function

    Set ResolveMasterVariables = LLVariables.Create(resolvedDictionary)
End Function

'The variables manager wraps the first table of the Variables sheet, which is
'the one MasterSetupPreparation maintains.
Public Function ResolveMasterSetupVariables(Optional ByVal hostSheet As Worksheet) As MasterSetupVariables
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterVariablesSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function
    If targetSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveMasterSetupVariables = MasterSetupVariables.Create(targetSheet.ListObjects(1))
End Function

Public Function ResolveMasterDropdowns(Optional ByVal hostSheet As Worksheet, _
                                       Optional ByVal headerPrefix As String = DEFAULT_DROPDOWN_PREFIX) As DropdownLists
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterDropdownsSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterDropdowns = DropdownLists.Create(targetSheet, headerPrefix)
End Function

Public Function ResolveMasterPasswords(Optional ByVal hostSheet As Worksheet) As Passwords
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterPasswordsSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterPasswords = Passwords.Create(targetSheet)
End Function

'The master registry on "__updated" is hand-built. Nothing in master setup calls
'AddSheet or AddColumns, so the watcher is used here to read flags and to reset
'them, never to build the tables. UpdatedValues.SwitchTags walks every table on
'the sheet and resets any that carries an "updated" column, which is how the
'master tables are reached even though this class never created them.
Public Function ResolveMasterUpdatedValues(Optional ByVal registrySheet As Worksheet) As UpdatedValues
    Dim targetSheet As Worksheet

    If registrySheet Is Nothing Then
        Set targetSheet = ResolveMasterRegistrySheet()
    Else
        Set targetSheet = registrySheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterUpdatedValues = UpdatedValues.Create(targetSheet)
End Function

Public Function ResolveMasterDevelopment(Optional ByVal devSheet As Worksheet, _
                                         Optional ByVal codeSheet As Worksheet) As Development
    Dim targetDevSheet As Worksheet
    Dim targetCodeSheet As Worksheet

    If devSheet Is Nothing Then
        Set targetDevSheet = ResolveMasterDevelopmentSheet()
    Else
        Set targetDevSheet = devSheet
    End If

    If codeSheet Is Nothing Then
        Set targetCodeSheet = Nothing
    Else
        Set targetCodeSheet = codeSheet
    End If

    If targetDevSheet Is Nothing Then Exit Function

    If targetCodeSheet Is Nothing Then
        Set ResolveMasterDevelopment = Development.Create(targetDevSheet)
    Else
        Set ResolveMasterDevelopment = Development.Create(targetDevSheet, targetCodeSheet)
    End If
End Function

'@section Tables Management utilities
'===============================================================================
'@sub-title Add a batch of rows to every table of a sheet, or trim their empty rows.
'@details Every sheet gains DEFAULT_ROW_BATCH rows per add. A trim hands
'CustomTable a threshold of 0, so it counts the formula columns of the first
'row: a disease line carrying only its label and choice values formulas is
'empty, a choices line carrying only its label formula is empty, and a
'variables line is empty when nothing is typed. The first row of a table
'always stays. The sheet kind (ResolveMasterSheetKind) decides the protection
'put back at the end. A failure is raised at the caller once the sheet is
'protected again and the application state is restored, so the ribbon can
'say what went wrong.
Public Sub ManageRows(ByVal targetSheet As Worksheet, ByVal addRows As Boolean)

    Dim lo As ListObject
    Dim wrapper As CustomTable
    Dim sheetKind As String
    Dim scope As ApplicationState
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If targetSheet Is Nothing Then Exit Sub

    sheetKind = ResolveMasterSheetKind(targetSheet)

    'The handler is armed ABOVE ApplyBusyState. That call writes screen
    'updating and alerts before the settings that can refuse, so a raise
    'there leaves the screen off and has to reach Cleanup.
    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    UnProtectMasterSetupSheet targetSheet

    For Each lo In targetSheet.ListObjects
        Set wrapper = CustomTable.Create(lo)
        If addRows Then
            wrapper.AddRows nbRows:=DEFAULT_ROW_BATCH
        Else
            wrapper.RemoveRows totalCount:=EMPTY_ROW_THRESHOLD
        End If
        'A disease table grown through Resize carries neither format nor
        'formula on its new rows; the frame and the two line formulas go
        'back over the whole table.
        If addRows And sheetKind = DISEASE_TAG_VALUE Then
            DiseaseSheet.FrameTable lo
            DiseaseSheet.WriteLineFormulas lo
        End If
    Next lo

    ProtectMasterSetupSheet targetSheet, sheetKind

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore would
    'come straight back to this label and raise again. The failure path puts
    'the sheet back under protection before the state is restored.
    On Error Resume Next
    If errNumber <> 0 Then ProtectMasterSetupSheet targetSheet, sheetKind
    If Not scope Is Nothing Then scope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Sub

Handler:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Debug.Print "ManageRows - "; targetSheet.Name; " addRows: "; addRows; " error "; errNumber; " "; errDescription
    Resume Cleanup
End Sub


'@sub-title Show every filtered row of the tables of a sheet.
'@details The sheet kind (ResolveMasterSheetKind) decides the protection
'put back at the end. A failure is raised at the caller once the sheet is
'protected again and the application state is restored, the way ManageRows
'does it, so the ribbon can say what went wrong.
Public Sub ClearMasterSheetFilters(ByVal targetSheet As Worksheet)

    Dim lo As ListObject
    Dim sheetKind As String
    Dim scope As ApplicationState
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If targetSheet Is Nothing Then Exit Sub

    sheetKind = ResolveMasterSheetKind(targetSheet)

    'The handler is armed ABOVE ApplyBusyState. That call writes screen
    'updating and alerts before the settings that can refuse, so a raise
    'there leaves the screen off and has to reach Cleanup.
    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    UnProtectMasterSetupSheet targetSheet

    For Each lo In targetSheet.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            'ShowAllData raises on a table with nothing filtered; the
            'table is already in the state wanted.
            On Error Resume Next
                lo.AutoFilter.ShowAllData
            On Error GoTo Handler
        End If
    Next lo

    If targetSheet.AutoFilterMode Then
        targetSheet.AutoFilterMode = False
    End If

    ProtectMasterSetupSheet targetSheet, sheetKind

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore would
    'come straight back to this label and raise again. The failure path puts
    'the sheet back under protection before the state is restored.
    On Error Resume Next
    If errNumber <> 0 Then ProtectMasterSetupSheet targetSheet, sheetKind
    If Not scope Is Nothing Then scope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Sub

Handler:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Debug.Print "Clear filters - "; targetSheet.Name; " error "; errNumber; " "; errDescription
    Resume Cleanup
End Sub

Public Sub UnProtectMasterSetupSheet(ByVal targetSheet As Worksheet)
    Dim passManager As Passwords

    If targetSheet Is Nothing Then Exit Sub

    Set passManager = ResolveMasterPasswords()
    If passManager Is Nothing Then Exit Sub

    passManager.UnProtect targetSheet.Name
End Sub

Public Sub ProtectMasterSetupSheet(ByVal targetSheet As Worksheet, ByVal sheetTag As String)
    Dim passManager As Passwords
    Dim allowDelete As Boolean
    Dim normalized As String

    If targetSheet Is Nothing Then Exit Sub

    normalized = LCase$(Trim$(sheetTag))
    Select Case normalized
        Case "variable", "variables", "choices"
            allowDelete = True
        Case Else
            allowDelete = False
    End Select

    Set passManager = ResolveMasterPasswords()
    If passManager Is Nothing Then Exit Sub

    passManager.Protect targetSheet.Name, allowDeletingRows:=allowDelete
End Sub

Public Sub SortMasterVariablesTables(ByVal targetSheet As Worksheet)
    Dim table As ListObject
    Dim wrapper As CustomTable
    Dim columns As BetterArray

    If targetSheet Is Nothing Then Exit Sub

    For Each table In targetSheet.ListObjects
        Set wrapper = CustomTable.Create(table)
        Set columns = New BetterArray
        columns.LowerBound = 1
        columns.Push "Variable Section", "Variable Name"
        wrapper.Sort colName:="Variable Order", colList:=columns, directSort:=True, strictSearch:=False
        wrapper.Sort colName:=""
    Next table
End Sub

Public Sub ClearMasterSheetData(ByVal targetSheet As Worksheet)
    Dim table As ListObject
    Dim dataRange As Range

    If targetSheet Is Nothing Then Exit Sub

    For Each table In targetSheet.ListObjects
        Set dataRange = table.DataBodyRange
        If Not dataRange Is Nothing Then
            dataRange.ClearContents
        End If
    Next table
End Sub

Public Function ShouldManageMasterSheet(ByVal sheetName As String) As Boolean
    Dim dropdowns As DropdownLists
    Dim configSheets As BetterArray

    sheetName = Trim$(sheetName)
    If LenB(sheetName) = 0 Then Exit Function

    ShouldManageMasterSheet = True

    Set dropdowns = ResolveMasterDropdowns()
    If dropdowns Is Nothing Then Exit Function

    Set configSheets = dropdowns.Values(CONFIG_SHEETS_LIST)
    If configSheets Is Nothing Then Exit Function

    ShouldManageMasterSheet = Not ContainsValue(configSheets, sheetName)
End Function

'@sub-title Answer the kind of a master setup sheet: disease, variables, choices, or empty.
'@details The hidden sheetTag is read first; the disease builder and the
'preparation write it. A sheet with no tag is known by its name, so the
'Variables and Choices sheets answer their kind before the preparation has
'run on them, and the ribbon row buttons work on a file opened with the
'events off.
Public Function ResolveMasterSheetKind(ByVal targetSheet As Worksheet) As String
    Dim store As HiddenNames
    Dim tagValue As String

    If targetSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set store = HiddenNames.Create(targetSheet)
    On Error GoTo 0
    If Not store Is Nothing Then
        tagValue = LCase$(Trim$(store.ValueAsString(SHEET_TAG_NAME)))
    End If

    Select Case tagValue
        Case "disease", "dis"
            ResolveMasterSheetKind = DISEASE_TAG_VALUE
        Case "variables", "variable", "var"
            ResolveMasterSheetKind = VARIABLES_TAG_VALUE
        Case "choices", "choice", "choi"
            ResolveMasterSheetKind = CHOICES_TAG_VALUE
        Case Else
            If StrComp(targetSheet.Name, VARIABLES_SHEETNAME, vbTextCompare) = 0 Then
                ResolveMasterSheetKind = VARIABLES_TAG_VALUE
            ElseIf StrComp(targetSheet.Name, CHOICES_SHEETNAME, vbTextCompare) = 0 Then
                ResolveMasterSheetKind = CHOICES_TAG_VALUE
            End If
    End Select
End Function

'A disease sheet is recognised by its hidden sheetTag (owner rule,
'2026-08-25). The old master setup used a marker cell; the cell is gone.
Public Function IsMasterDiseaseSheet(ByVal targetSheet As Worksheet) As Boolean
    Dim store As HiddenNames

    If targetSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set store = HiddenNames.Create(targetSheet)
    On Error GoTo 0
    If store Is Nothing Then Exit Function

    IsMasterDiseaseSheet = (StrComp(Trim$(store.ValueAsString(SHEET_TAG_NAME)), _
                                    DISEASE_TAG_VALUE, vbTextCompare) = 0)
End Function

Public Function ResolveNextDiseaseIndex(Optional ByVal targetBook As Workbook) As Long
    Dim sh As Worksheet
    Dim count As Long

    Set targetBook = ResolveMasterSetupWorkbook(targetBook)

    For Each sh In targetBook.Worksheets
        If IsMasterDiseaseSheet(sh) Then count = count + 1
    Next sh

    ResolveNextDiseaseIndex = count + 1
End Function

Public Function CleanMasterSheetName(ByVal rawName As String) As String
    Const MAX_LENGTH As Long = 31
    Dim sanitized As String

    sanitized = ReplaceInvalidWorksheetChars(Trim$(rawName))
    If Len(sanitized) > MAX_LENGTH Then
        sanitized = Left$(sanitized, MAX_LENGTH)
    End If

    CleanMasterSheetName = sanitized
End Function

Public Function SafeValue(ByVal candidate As Variant) As String
    If IsError(candidate) Then
        SafeValue = vbNullString
    Else
        SafeValue = CStr(candidate)
    End If
End Function

'@section Private helpers
'===============================================================================
Private Function ReplaceInvalidWorksheetChars(ByVal valueText As String) As String
    Dim cleaned As String

    cleaned = Application.WorksheetFunction.Substitute(valueText, Chr$(160), " ")
    cleaned = Application.WorksheetFunction.Clean(cleaned)

    cleaned = Replace(cleaned, "<", "_")
    cleaned = Replace(cleaned, ">", "_")
    cleaned = Replace(cleaned, ":", "_")
    cleaned = Replace(cleaned, "|", "_")
    cleaned = Replace(cleaned, "?", "_")
    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "\", "_")
    cleaned = Replace(cleaned, "[", "_")
    cleaned = Replace(cleaned, "]", "_")
    cleaned = Replace(cleaned, "*", "_")
    cleaned = Replace(cleaned, ".", "_")
    cleaned = Replace(cleaned, """", "_")

    ReplaceInvalidWorksheetChars = cleaned
End Function

Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

Private Function ContainsValue(ByVal items As BetterArray, ByVal expected As String) As Boolean
    Dim idx As Long
    Dim candidate As Variant

    If items Is Nothing Then Exit Function
    If items.Length = 0 Then Exit Function

    For idx = items.LowerBound To items.UpperBound
        candidate = items.Item(idx)
        If NormalizeText(CStr(candidate)) = NormalizeText(expected) Then
            ContainsValue = True
            Exit Function
        End If
    Next idx
End Function
