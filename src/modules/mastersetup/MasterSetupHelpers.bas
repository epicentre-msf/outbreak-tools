Attribute VB_Name = "MasterSetupHelpers"
Option Explicit

'@Folder("Master Setup")
'@ModuleDescription("Utility helpers shared across master setup modules.")
'@depends DropdownLists, IDropdownLists, CustomTable, ICustomTable, Passwords, IPasswords, Translation, ITranslationObject, BetterArray

Private Const VARIABLES_SHEETNAME As String = "Variables"
Private Const TRANSLATIONS_SHEETNAME As String = "Translations"
Private Const CHOICES_SHEETNAME As String = "Choices"
Private Const DROPDOWNS_SHEETNAME As String = "__dropdowns"
Private Const REGISTRY_SHEETNAME As String = "__updated"
Private Const PASSWORDS_SHEETNAME As String = "__pass"
Private Const DEVELOPMENT_SHEETNAME As String = "Dev"
Private Const CONFIG_SHEETS_LIST As String = "__configSheets"
Private Const RIBBON_TRANSLATION As String = "__ribbonTranslation"

Private Const START_ROW_VARIABLES As Long = 5
Private Const START_COLUMN_VARIABLES As Long = 1
Private Const START_ROW_CHOICES As Long = 4
Private Const START_COLUMN_CHOICES As Long = 1
Private Const DEFAULT_DROPDOWN_PREFIX As String = "dropdown_"
Private Const DISEASE_MARKER_VALUE As String = "DISSHEET"
Private Const DISEASE_MARKER_ROW As Long = 2
Private Const DISEASE_MARKER_COLUMN As Long = 4
Private Const DEFAULT_ROW_BATCH As Long = 5
Private Const DEFAULT_ROW_TARGET As Long = 1

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
        Case "ribbontrads", "ribtrads", "ribtrad"
            ResolveMasterSetupSheetName = RIBBON_TRANSLATION
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
Public Function ResolveMasterDictionary(Optional ByVal hostSheet As Worksheet) As ILLdictionary
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterVariablesSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterDictionary = LLdictionary.Create(targetSheet, START_ROW_VARIABLES, START_COLUMN_VARIABLES)
End Function

Public Function ResolveMasterChoices(Optional ByVal hostSheet As Worksheet) As ILLChoices
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterChoicesSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterChoices = LLChoices.Create(targetSheet, START_ROW_CHOICES, START_COLUMN_CHOICES)
End Function

Public Function ResolveMasterVariables(Optional ByVal dictionary As ILLdictionary, _
                                       Optional ByVal hostSheet As Worksheet) As ILLVariables
    Dim resolvedDictionary As ILLdictionary

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

Public Function ResolveMasterDropdowns(Optional ByVal hostSheet As Worksheet, _
                                       Optional ByVal headerPrefix As String = DEFAULT_DROPDOWN_PREFIX) As IDropdownLists
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterDropdownsSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterDropdowns = DropdownLists.Create(targetSheet, headerPrefix)
End Function

Public Function ResolveMasterPasswords(Optional ByVal hostSheet As Worksheet) As IPasswords
    Dim targetSheet As Worksheet

    If hostSheet Is Nothing Then
        Set targetSheet = ResolveMasterPasswordsSheet()
    Else
        Set targetSheet = hostSheet
    End If

    If targetSheet Is Nothing Then Exit Function

    Set ResolveMasterPasswords = Passwords.Create(targetSheet)
End Function

Public Function ResolveMasterUpdatedValues(Optional ByVal registrySheet As Worksheet) As IUpdatedValues
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
                                         Optional ByVal codeSheet As Worksheet) As IDevelopment
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
Public Sub ManageRows(ByVal targetSheet As Worksheet, ByVal addRows As Boolean)

    Dim lo As ListObject
    Dim wrapper As ICustomTable
    Dim store As IHiddenNames
    Dim sheetTag As String
    Dim rowCount As Long
    Dim scope As IApplicationState

    If targetSheet Is Nothing Then Exit Sub

    Set store = HiddenNames.Create(targetSheet)
    If Not store Is Nothing Then
        sheetTag = store.ValueAsString("sheetTag")
    End If

    Select Case LCase$(Trim$(sheetTag))
        Case "disease", "dis"
            rowCount = IIf(addRows, 2, 10)
        Case "choices", "choi"
            rowCount = IIf(addRows, 0, 5)
        Case "variables", "variable", "var"
            rowCount = IIf(addRows, 1, 20)
        Case Else
            rowCount = IIf(addRows, 0, 10)
    End Select

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    On Error GoTo Handler

    UnProtectMasterSetupSheet targetSheet, sheetTag

    For Each lo In targetSheet.ListObjects
        Set wrapper = CustomTable.Create(lo)
        If addRows Then
            wrapper.AddRows nbRows:=rowCount
        Else
            wrapper.RemoveRows totalCount:=rowCount
        End If
    Next lo

    ProtectMasterSetupSheet targetSheet, sheetTag

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "ManageRows - "; targetSheet.Name; " addRows: "; addRows; " error "; Err.Number; " "; Err.Description
    Resume Cleanup
End Sub


Public Sub ClearMasterSheetFilters(ByVal targetSheet As Worksheet)

    Dim lo As ListObject
    Dim scope As IApplicationState

    If targetSheet Is Nothing Then Exit Sub
    
    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False


    UnProtectMasterSetupSheet targetSheet, vbNullString

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

    ProtectMasterSetupSheet targetSheet, vbNullString

Cleanup:
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "Clear filters - "; targetSheet.Name; " addRows: "; addRows; " error "; Err.Number; " "; Err.Description
    Resume Cleanup
End Sub

Public Sub UnProtectMasterSetupSheet(ByVal targetSheet As Worksheet, ByVal sheetTag As String)
    Dim passwords As IPasswords

    If targetSheet Is Nothing Then Exit Sub

    Set passwords = ResolveMasterPasswords()
    If passwords Is Nothing Then Exit Sub

    passwords.UnProtect targetSheet.Name
End Sub

Public Sub ProtectMasterSetupSheet(ByVal targetSheet As Worksheet, ByVal sheetTag As String)
    Dim passwords As IPasswords
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

    Set passwords = ResolveMasterPasswords()
    If passwords Is Nothing Then Exit Sub

    passwords.Protect targetSheet.Name, allowDeletingRows:=allowDelete
End Sub

Public Sub SortMasterVariablesTables(ByVal targetSheet As Worksheet)
    Dim table As ListObject
    Dim wrapper As ICustomTable
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
    Dim dropdowns As IDropdownLists
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

Public Function IsMasterDiseaseSheet(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function
    IsMasterDiseaseSheet = (StrComp(CStr(targetSheet.Cells(DISEASE_MARKER_ROW, DISEASE_MARKER_COLUMN).Value), _
                                    DISEASE_MARKER_VALUE, vbTextCompare) = 0)
End Function

Public Function ResolveRibbonTranslations(Optional ByVal workbook As Workbook) As ITranslationObject
    Dim tagSheet As Worksheet
    Dim table As ListObject
    Dim languageTag As String
    Dim targetBook As Workbook

    Set targetBook = ResolveMasterSetupWorkbook(workbook)

    On Error Resume Next
        Set tagSheet = targetBook.Worksheets(RIBBON_TRANSLATION_SHEET)
        Set table = tagSheet.ListObjects(RIBBON_TRANSLATION_TABLE)
    On Error GoTo 0

    If table Is Nothing Then Exit Function

    languageTag = SafeValue(tagSheet.Range(RIBBON_LANGUAGE_RANGE).Value)
    Set ResolveRibbonTranslations = Translation.Create(table, languageTag)
End Function

Public Function ResolveRibbonLanguageTag(Optional ByVal workbook As Workbook) As String
    Dim tagSheet As Worksheet
    Dim targetBook As Workbook

    Set targetBook = ResolveMasterSetupWorkbook(workbook)

    On Error Resume Next
        Set tagSheet = targetBook.Worksheets(RIBBON_TRANSLATION_SHEET)
    On Error GoTo 0

    If tagSheet Is Nothing Then
        ResolveRibbonLanguageTag = vbNullString
    Else
        ResolveRibbonLanguageTag = SafeValue(tagSheet.Range(RIBBON_LANGUAGE_RANGE).Value)
    End If
End Function

Public Function TranslateValue(ByVal translations As ITranslationObject, _
                               ByVal key As String, _
                               ByVal fallback As String) As String
    If translations Is Nothing Then
        TranslateValue = fallback
    ElseIf translations.ValueExists(key) Then
        TranslateValue = translations.TranslatedValue(key)
    Else
        TranslateValue = fallback
    End If
End Function

Public Function ResolveNextDiseaseIndex(Optional ByVal workbook As Workbook) As Long
    Dim targetBook As Workbook
    Dim sh As Worksheet
    Dim count As Long

    Set targetBook = ResolveMasterSetupWorkbook(workbook)

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
