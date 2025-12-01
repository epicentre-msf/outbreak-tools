VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analyses setup sheets and publishes consistency reports"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Analyses setup sheets and publishes consistency reports")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends Analysis, ApplicationState, BetterArray, Checking, CheckingOutput, CustomTable, DropdownLists, FormulaData, Formulas, IAnalysis, ISetupErrors, IDropdownLists, LLChoices, LLVariables, LLdictionary, Passwords, ProjectError, TablesSpecs

Option Explicit

Implements ISetupErrors

'@section State
'===============================================================================
Private Type TSetupErrorsState
    hostBook As Workbook              'Target workbook containing setup data
    formulaSheet As Worksheet         'Worksheet storing helper tables and formulas
    passwordService As IPasswords     'Password handler protecting/unprotecting sheets
    dictionary As ILLdictionary       'Dictionary accessor for setup data
    choices As ILLChoices             'Choices accessor for list validation
    formulaStore As IFormulaData      'Formula data repository used for validation
    checkTables As BetterArray        'Collection of generated checkings
    appScope As IApplicationState     'Application state guard ensuring UI stays responsive
    calcBeforeSave As Boolean         'Original CalculateBeforeSave value to restore
    calcBeforeSaveCaptured As Boolean 'Flag indicating CalculateBeforeSave snapshot is stored
End Type

Private this As TSetupErrorsState

Private Const CLASS_NAME As String = "SetupErrors"

Private Const SHEET_FORMULA As String = "__formula"
Private Const SHEET_PASSWORDS As String = "__pass"
Private Const SHEET_CHECK_OUTPUT As String = "__checkRep"
Private Const SHEET_DROPDOWN As String = "__variables"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const SHEET_REGISTRY As String = "__updated"

'@section List of errors to take in account
'===============================================================================

Private Const ANA_ADGR_FLIP As String               = "Analysis, line %1--There is no graph for this table, so ""flip coordinates"" will be ignored."
Private Const ANA_EMPTY_TAB As String               = "Analysis, line %1--This line is completely empty, the table will be ignored in analysis."
Private Const ANA_EMPTY_TITLE As String             = "Analysis, line %1--The table does not have a title."
Private Const ANA_INCOR_FORM As String              = "Analysis, line %1--The formula specified for the table will eventually fail and not be parsed. %3"
Private Const ANA_INV_TAB As String                 = "Analysis, line %1--The table is invalid. There is an error in specifications: %2"
Private Const ANA_TS_NEWSEC As String               = "Analysis, line %1--This table is invalid and is a new section of time series. This will likely generate errors or impact all the time series sheet."
Private Const ANA_UAFLIP_PERC As String             = "Analysis, line %1--The graph for this univariate table will not contain percentages because ""flip coordinates"" = yes."
Private Const CHOI_EMPTYCHOI_LAB As String          = "Choices, line %1--There is no attached ""List Name"" value for Label ""%2""."
Private Const CHOI_EMPTYCHOI_ORDER As String        = "Choices, line %1--There is no attached ""List Name"" value for Ordering list ""%2""."
Private Const CHOI_EMPTY_ORDER As String            = "Choices, line %1--There is no value of ""Ordering list"" for choice ""%2""."
Private Const CHOI_MIS_LAB As String                = "Choices, line %1--There is a missing Label for choice ""%2""."
Private Const CHOI_UNFOUND_CHOI As String           = "Choices, ""List Name"" column -- Choice name ""%1"" is declared in choices sheet but never used."
Private Const DICT_CAT_NOTFOUND As String           = "Dictionary, line %1--in CHOICE_FORMULA formula, category ""%2"" does not exists for choice ""%3"". Consider adding it to the choice sheet."
Private Const DICT_CHOI_EMPTY As String             = "Dictionary, line %1--choice manual ""%2"" is not present in choices sheet, please add its categories in the Choices sheet."
Private Const DICT_CHOIFORM_EMPTY As String         = "Dictionary, line %1--choice formula ""%2"" is not present in Choices sheet, please add its categories in the choice sheet."
Private Const DICT_EMPTY_SHEET As String            = "Dictionary, line %1--Sheet names should not be empty, please provide a good sheet name for variable ""%2"" at line %1."
Private Const DICT_EXPORT_NA As String              = "Dictionary, exports column--""Export %1"" column is filled in the dictionary, but ""Export Number"" %1 status is not active in the export sheet."
Private Const DICT_FORMAT_CONTROL As String         = "Dictionary, line %1--Format requires the type of the variable, please consider adding the type for variable %2."
Private Const DICT_INCOR_FORM As String             = "Dictionary, line %1--The formula for variable ""%2"" will eventually fail and not be parsed. %3."
Private Const DICT_INCOR_MAX As String              = "Dictionary, line %1--The Max validation for variable ""%2"" will eventually fail and not be parsed. %3."
Private Const DICT_INCOR_MIN As String              = "Dictionary, line %1--The Min validation for variable ""%2"" will eventually fail and not be parsed. %3."
Private Const DICT_MAIN_LAB As String               = "Dictionary, line %1--The main label of variable ""%2"" is empty."
Private Const DICT_UNIQUE_VAL As String             = "Dictionary, line %1--Duplicate values of variable %2 will be highlighted in red because Unique = ""yes""."
Private Const DICT_UNKNOWN_CONTROL As String        = "Dictionary, line %1--The control ""%2"" of variable ""%3"" is unknown."
Private Const DICT_VALID_CONTROL As String          = "Dictionary, line %1--Validation requires the type of the variable, please consider adding the type for variable %2."
Private Const DICT_VAR_LENGTH As String             = "Dictionary, line %1--Variable ""%2"" length is less than 4. Variable should have length >= 4."
Private Const DICT_VAR_UNIQUE As String             = "Dictionary, line %1--Variable ""%2"" is duplicate, variable names should be unique."
Private Const EXP_ACT_EMPTY As String               = "Exports, line %1--The export is active, but the corresponding column is empty in the dictionary. You may end up with empty data in this export. Please correct the status of export %2."
Private Const EXP_FILENAME_VARNOTVLIST As String    = "Exports, line %1--The Export Number %2 file name contains a variable named ""%3"" which is not a vlist1D variable, but a hlist2D. It will be considered as text in exports."
Private Const EXP_FILENAME_VARUNFOUND As String     = "Exports, line %1--The Export Number %2 file name contains a variable named ""%3"" which does not exists. It will be considered as text in exports."
Private Const EXP_ID_PASSWD As String               = "Exports, line %1--The Export Number %2 has no password, but keeps identifiers."
Private Const EXP_MIS_FORM As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""File Format"" (xlsx/xlsb) attached to it."
Private Const EXP_MIS_HEAD As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Header format"" attached to it. It will default to ""variables names""."
Private Const EXP_MIS_LAB As String                 = "Exports, line %1--The Export Number %2 is active, but there is no label attached to it."
Private Const EXP_MIS_META As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Export Metadata"" (yes/no) attached to it."
Private Const EXP_MIS_NAME As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""File name"" attached to it."
Private Const EXP_MIS_PASS As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Password"" (yes/no) attached to it."
Private Const EXP_MIS_TRAD As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Export Translation"" (yes/no) attached to it."
Private Const EXP_UNFOUND_DICTCOLUMN As String      = "Exports, line %1--The Export Number %2 is active, but the corresponding column in the dictionary is missing."


'@section Factory
'===============================================================================
'@sub-title Create and initialise a setup checker instance.
'@param hostBook Workbook containing setup data.
'@return ISetupErrors ready to run checks.
Public Function Create(ByVal hostBook As Workbook, _
                       Optional ByVal passwordService As IPasswords, _
                       Optional ByVal dictionary As ILLdictionary, _
                       Optional ByVal choices As ILLChoices, _
                       Optional ByVal formulaStore As IFormulaData, _
                       Optional ByVal stateScope As IApplicationState) As ISetupErrors
    Dim instance As SetupErrors

    Set instance = New SetupErrors
    instance.Initialise hostBook, passwordService, dictionary, choices, formulaStore, stateScope

    Set Create = instance.Self
End Function

'@section Initialisation
'===============================================================================
'@sub-title Bind the checker to required collaborators.
Public Sub Initialise(ByVal hostBook As Workbook, _
                      Optional ByVal passwordService As IPasswords, _
                      Optional ByVal dictionary As ILLdictionary, _
                      Optional ByVal choices As ILLChoices, _
                      Optional ByVal formulaStore As IFormulaData, _
                      Optional ByVal stateScope As IApplicationState)

    ValidateWorkbook hostBook
    Set this.hostBook = hostBook

    Set this.passwordService = passwordService
    Set this.dictionary = dictionary
    Set this.choices = choices
    Set this.formulaStore = formulaStore
    Set this.appScope = stateScope

    ConfigureDependencies
    EnsureCheckContainer
End Sub

'@section Execution
'===============================================================================
'@sub-title Execute all setup checks and publish results.
Public Sub Run()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    EnsureInitialised

    EnterBusyState
    On Error GoTo RunFailed

        ResetCheckContainer
        CheckDictionary
        CheckChoice
        CheckExports
        'CheckAnalysis
        CheckTranslations
        PrintReport

    LeaveBusyState
    Exit Sub

RunFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@section Results
'===============================================================================
'@sub-title Expose generated checkings for inspection.
Public Property Get Checkings() As BetterArray
    EnsureInitialised
    Set Checkings = this.checkTables
End Property

'@section Helper workflow
'===============================================================================
'@sub-title Self reference used by factory helpers.
Public Property Get Self() As ISetupErrors
    Set Self = Me
End Property

'@sub-title Ensure the checker has been initialised.
Private Sub EnsureInitialised()
    If (this.hostBook Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupErrors.Initialise must be called before usage."
    End If
End Sub

'@sub-title Prepare collaborators when not provided.
Private Sub ConfigureDependencies()
    Dim formulaSheet As Worksheet

    If Not this.formulaSheet Is Nothing Then
        Set formulaSheet = this.formulaSheet
    ElseIf (this.formulaStore Is Nothing) Then
        Set formulaSheet = ResolveWorksheet(SHEET_FORMULA)
        Set this.formulaSheet = formulaSheet
    End If

    If this.passwordService Is Nothing Then
        Set this.passwordService = Passwords.Create(ResolveWorksheet(SHEET_PASSWORDS))
    End If

    If this.dictionary Is Nothing Then
        Set this.dictionary = LLdictionary.Create(ResolveWorksheet(SHEET_DICTIONARY), 5, 1)
    End If

    If this.choices Is Nothing Then
        Set this.choices = LLChoices.Create(ResolveWorksheet(SHEET_CHOICES), 4, 1)
    End If

    If this.formulaStore Is Nothing Then
        If formulaSheet Is Nothing Then
            Set formulaSheet = ResolveWorksheet(SHEET_FORMULA)
            Set this.formulaSheet = formulaSheet
        End If
        Set this.formulaStore = FormulaData.Create(formulaSheet)
    End If

End Sub

'@sub-title Ensure the internal container for checkings exists.
Private Sub EnsureCheckContainer()
    If this.checkTables Is Nothing Then
        Set this.checkTables = New BetterArray
    Else
        ResetCheckContainer
    End If
End Sub

'@sub-title Reset the container prior to running checks.
Private Sub ResetCheckContainer()
    If this.checkTables Is Nothing Then
        Set this.checkTables = New BetterArray
    Else
        this.checkTables.Clear
    End If
End Sub

'@sub-title Retrieve or create the application state scope.
Private Function ApplicationScope() As IApplicationState
    If this.appScope Is Nothing Then
        Set this.appScope = ApplicationState.Create(Application)
    End If

    Set ApplicationScope = this.appScope
End Function

'@sub-title Suspend Excel UI features while running checks.
Private Sub EnterBusyState()
    Dim scope As IApplicationState

    Set scope = ApplicationScope()

    If Not this.calcBeforeSaveCaptured Then
        this.calcBeforeSave = Application.CalculateBeforeSave
        this.calcBeforeSaveCaptured = True
    End If

    Application.CalculateBeforeSave = False

    If Not scope.IsBusy Then
        scope.ApplyBusyState, True
    End If
End Sub

'@sub-title Restore Excel UI configuration.
Private Sub LeaveBusyState(Optional ByVal silent As Boolean = False)
    Dim scope As IApplicationState

    If this.calcBeforeSaveCaptured Then
        On Error Resume Next
            Application.CalculateBeforeSave = this.calcBeforeSave
        On Error GoTo 0
        this.calcBeforeSaveCaptured = False
    End If

    Set scope = this.appScope
    If scope Is Nothing Then Exit Sub

    On Error GoTo RestoreFailed
        scope.Restore
    On Error GoTo 0
    Exit Sub

RestoreFailed:
    If Not silent Then
        Err.Raise Err.Number, CLASS_NAME, "Failed to restore application state: " & Err.Description
    Else
        Err.Clear
    End If
End Sub

'@sub-title Resolve a workbook worksheet by name.
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet

    On Error Resume Next
        Set target = this.hostBook.Worksheets(sheetName)
    On Error GoTo 0

    If target Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Worksheet '" & sheetName & "' not found in setup workbook."
    End If

    Set ResolveWorksheet = target
End Function

'@sub-title Validate the incoming workbook reference.
Private Sub ValidateWorkbook(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupErrors requires a workbook reference."
    End If
End Sub

'@section Core checking routines
'===============================================================================
'@sub-title Validate dictionary sheet for structure, duplicates, and formulas.
Private Sub CheckDictionary()

    Dim check As IChecking
    Dim csTab As ICustomTable
    Dim expTab As ICustomTable
    Dim varRng As Range
    Dim sheetRng As Range
    Dim workFun As WorksheetFunction
    Dim varValue As String
    Dim sheetValue As String
    Dim shDictionary As Worksheet
    Dim shExports As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim cellRng As Range
    Dim controlDetailsValue As String
    Dim controlValue As String
    Dim setupForm As Object
    Dim checkingCounter As Long
    Dim choiCategories As BetterArray
    Dim formCategories As BetterArray
    Dim controlsList As BetterArray
    Dim choiName As String
    Dim tabCounter As Long
    Dim catValue As String
    Dim expCounter As Long
    Dim expRng As Range
    Dim expStatusRng As Range
    Dim minValue As String
    Dim maxValue As String
    Dim mainVarRng As Range
    Dim mainLabValue As String
    Dim uniqueValue As String
    Dim typeValue As String
    Dim formatValue As String

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo DictionaryFailed

    Set shDictionary = ResolveWorksheet(SHEET_DICTIONARY)
    Set shExports = ResolveWorksheet(SHEET_EXPORTS)
    Set check = Checking.Create("Dictionary incoherences", "Type--Where?--Details")
    Set csTab = CustomTable.Create(shDictionary.ListObjects(1), idCol:="Variable Name")
    Set expTab = CustomTable.Create(shExports.ListObjects(1), idCol:="Export Number")
    Set workFun = Application.WorksheetFunction

    Set choiCategories = New BetterArray
    Set formCategories = New BetterArray
    Set controlsList = New BetterArray

    this.passwordService.UnProtect SHEET_DICTIONARY
    this.passwordService.UnProtect SHEET_EXPORTS

    csTab.RemoveRows includeIds:=False
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name")
    Set sheetRng = csTab.DataRange("Sheet Name")
    Set mainVarRng = csTab.DataRange("Main Label")
    Set cellRng = varRng.Cells(varRng.Rows.Count, 1)

    controlsList.Push "choice_manual", "choice_formula", "formula", _
                      "geo", "hf", "custom", "list_auto", "case_when", _
                      "choice_custom", "choice_multiple"

    Do While cellRng.Row >= varRng.Row
        checkingCounter = 0
        varValue = workFun.Trim(cellRng.Value)

        If workFun.CountIf(varRng, varValue) > 1 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-unique"
            infoMessage = FormatMessage(DICT_VAR_UNIQUE, cellRng.Row, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        If Len(varValue) < 4 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-length"
            infoMessage = FormatMessage(DICT_VAR_LENGTH, cellRng.Row, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        mainLabValue = shDictionary.Cells(cellRng.Row, mainVarRng.Column)
        If mainLabValue = vbNullString Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-main-lab"
            infoMessage = FormatMessage(DICT_MAIN_LAB, cellRng.Row, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        sheetValue = shDictionary.Cells(cellRng.Row, sheetRng.Column)
        If sheetValue = vbNullString Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-empty-sheet"
            infoMessage = FormatMessage(DICT_EMPTY_SHEET, cellRng.Row, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        controlValue = csTab.Value("Control", varValue)
        controlDetailsValue = csTab.Value("Control Details", varValue)

        If (Not controlsList.Includes(controlValue)) And _
           (controlValue <> vbNullString) And _
           (Not (InStr(1, controlValue, "choice_multiple") = 1)) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unknown-control"
            infoMessage = FormatMessage(DICT_UNKNOWN_CONTROL, cellRng.Row, controlValue, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
        End If

        If (controlValue = "choice_formula") And (InStr(1, controlDetailsValue, "CHOICE_FORMULA") = 1) Then
            Set setupForm = ChoiceFormula.Create(controlDetailsValue)
            choiName = setupForm.choiceName()
            If Not this.choices.ChoiceExists(choiName) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choiform-empty"
                infoMessage = FormatMessage(DICT_CHOIFORM_EMPTY, cellRng.Row, choiName)
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            Else
                keyName = "dict-cat-notfound"
                Set choiCategories = this.choices.Categories(choiName)
                Set formCategories = setupForm.Categories()

                For tabCounter = formCategories.LowerBound To formCategories.UpperBound
                    catValue = formCategories.Item(tabCounter)
                    If Not choiCategories.Includes(catValue) Then
                        checkingCounter = checkingCounter + 1
                        infoMessage = FormatMessage(DICT_CAT_NOTFOUND, cellRng.Row, catValue, choiName)
                        check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingInfo
                    End If
                Next tabCounter
            End If
        End If

        If (controlValue = "formula" Or controlValue = "case_when" Or controlValue = "choice_formula") Then
            keyName = "dict-incor-form"
            infoMessage = FormulaMessage(controlDetailsValue, DICT_INCOR_FORM, cellRng.Row, varValue)
            If infoMessage <> vbNullString Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (controlValue = "choice_manual") Or _
           ((controlValue = "choice_custom") And (controlDetailsValue <> vbNullString)) Or _
           (InStr(1, controlValue, "choice_multiple") = 1) Then
            If Not this.choices.ChoiceExists(controlDetailsValue) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choi-empty"
                infoMessage = FormatMessage(DICT_CHOI_EMPTY, cellRng.Row, controlDetailsValue)
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        minValue = csTab.Value("Min", varValue)
        maxValue = csTab.Value("Max", varValue)

        If (minValue <> vbNullString) Then
            keyName = "dict-incor-min"
            infoMessage = FormulaMessage(minValue, DICT_INCOR_MIN, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (maxValue <> vbNullString) Then
            keyName = "dict-incor-max"
            infoMessage = FormulaMessage(maxValue, DICT_INCOR_MAX, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        typeValue = csTab.Value("Variable Type", varValue)
        If ((minValue <> vbNullString) Or (maxValue <> vbNullString)) And (typeValue = vbNullString) Then
            keyName = "dict-valid-control"
            infoMessage = FormatMessage(DICT_VALID_CONTROL, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingError
            End If
        End If

        formatValue = csTab.Value("Variable Format", varValue)
        If (formatValue <> vbNullString) And (typeValue = vbNullString) Then
            keyName = "dict-format-control"
            infoMessage = FormatMessage(DICT_FORMAT_CONTROL, cellRng.Row, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        uniqueValue = csTab.Value("Unique", varValue)
        If (uniqueValue = "yes") Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unique-val"
            infoMessage = FormatMessage(DICT_UNIQUE_VAL, cellRng.Row, varValue)
            check.Add keyName & "-row-" & cellRng.Row & "-" & checkingCounter, infoMessage, checkingInfo
        End If

        Set cellRng = cellRng.Offset(-1)
    Loop

    expTab.Sort "Export Number"

    For expCounter = 1 To this.dictionary.TotalNumberOfExports()
        Set expRng = csTab.DataRange("Export " & expCounter)
        Set expStatusRng = expTab.DataRange("Status")

        If (Not (expRng Is Nothing)) And (Not (expStatusRng Is Nothing)) Then
            If (workFun.CountBlank(expRng) <> expRng.Rows.Count) And (expStatusRng.Cells(expCounter, 1).Value <> "active") Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-export-na"
                infoMessage = FormatMessage(DICT_EXPORT_NA, CStr(expCounter))
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
            End If
        End If
    Next expCounter

    this.checkTables.Push check
    this.passwordService.Protect SHEET_DICTIONARY
    this.passwordService.Protect SHEET_EXPORTS
    On Error GoTo 0
    Exit Sub

DictionaryFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_DICTIONARY
        this.passwordService.Protect SHEET_EXPORTS
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Inspect choices sheet for unused or inconsistent entries.
Private Sub CheckChoice()
    Dim check As IChecking
    Dim shChoices As Worksheet
    Dim shDictionary As Worksheet
    Dim choiTab As ICustomTable
    Dim dictTab As ICustomTable
    Dim choiList As BetterArray
    Dim counter As Long
    Dim checkingCounter As Long
    Dim choiName As String
    Dim infoMessage As String
    Dim cellRng As Range
    Dim choiNameRng As Range
    Dim sortValue As String
    Dim choiLabValue As String
    Dim keyName As String
    Dim usedChoicesList As BetterArray
    Dim setupForm As Object
    Dim controlRng As Range
    Dim controlDetailsRng As Range
    Dim actualControl As String

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo ChoicesFailed

    Set shChoices = ResolveWorksheet(SHEET_CHOICES)
    Set shDictionary = ResolveWorksheet(SHEET_DICTIONARY)
    Set choiTab = CustomTable.Create(shChoices.ListObjects(1))
    Set usedChoicesList = New BetterArray

    this.passwordService.UnProtect SHEET_CHOICES

    this.choices.Sort
    choiTab.RemoveRows

    Set check = Checking.Create("Choices incoherences", "Type--Where?--Details")
    Set dictTab = CustomTable.Create(shDictionary.ListObjects(1))

    Set choiList = this.choices.AllChoices()
    Set choiNameRng = choiTab.DataRange("List Name")
    Set controlRng = dictTab.DataRange("Control")
    Set controlDetailsRng = dictTab.DataRange("Control Details")

    checkingCounter = 0

    For counter = 1 To controlDetailsRng.Rows.Count
        actualControl = controlRng.Cells(counter, 1).Value

        If (actualControl = "choice_manual") Or _
           (actualControl = "choice_custom") Or _
           (InStr(1, actualControl, "choice_multiple") = 1) Then

            If (controlDetailsRng.Cells(counter, 1).Value <> vbNullString) Then
                usedChoicesList.Push controlDetailsRng.Cells(counter, 1).Value
            End If
        ElseIf (actualControl = "choice_formula") Then
            If (controlDetailsRng.Cells(counter, 1).Value <> vbNullString) Then
                Set setupForm = ChoiceFormula.Create(controlDetailsRng.Cells(counter, 1).Value)
                choiName = setupForm.ChoiceName()
                If LenB(choiName) > 0 Then
                    usedChoicesList.Push choiName
                End If
            End If
        End If
    Next counter

    For counter = choiList.LowerBound To choiList.UpperBound
        choiName = CStr(choiList.Item(counter))
        If Not usedChoicesList.Includes(choiName) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-unfound-choi"
            infoMessage = FormatMessage(CHOI_UNFOUND_CHOI, choiName)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If
    Next counter

    Set cellRng = choiNameRng.Cells(choiNameRng.Rows.Count, 1)

    Do While cellRng.Row >= choiNameRng.Row
        choiName = cellRng.Value
        sortValue = cellRng.Offset(, 1).Value
        choiLabValue = cellRng.Offset(, 2).Value

        If (choiLabValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-emptychoi-lab"
            infoMessage = FormatMessage(CHOI_EMPTYCHOI_LAB, cellRng.Row, choiLabValue)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        If (sortValue <> vbNullString) And (choiName = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-emptychoi-order"
            infoMessage = FormatMessage(CHOI_EMPTYCHOI_ORDER, cellRng.Row, sortValue)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        If (sortValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-empty-order"
            infoMessage = FormatMessage(CHOI_EMPTY_ORDER, cellRng.Row, choiName)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If

        If (choiLabValue = vbNullString) And (choiName <> vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-mis-lab"
            infoMessage = FormatMessage(CHOI_MIS_LAB, cellRng.Row, choiName)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        End If

        Set cellRng = cellRng.Offset(-1)
    Loop

    this.checkTables.Push check
    this.passwordService.Protect SHEET_CHOICES
    On Error GoTo 0
    Exit Sub

ChoicesFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_CHOICES
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Audit exports sheet for missing metadata and dictionary alignment.
Private Sub CheckExports()

    Dim expTab As ICustomTable
    Dim counter As Long
    Dim shExports As Worksheet
    Dim check As IChecking
    Dim keyName As String
    Dim checkingCounter As Long
    Dim expStatus As String
    Dim cellRng As Range
    Dim infoMessage As String
    Dim keysList As BetterArray
    Dim headersList As BetterArray
    Dim messageList As BetterArray
    Dim headerCounter As Long
    Dim exportRng As Range
    Dim statusRng As Range
    Dim workFun As WorksheetFunction
    Dim fileNameList As BetterArray
    Dim actualFileName As String
    Dim fileCounter As Long
    Dim fileNameChunk As String
    Dim vars As ILLVariables
    Dim numberOfExports As Long
    Dim pwd As String
    Dim exportKeepsId As String

    Set shExports = ResolveWorksheet(SHEET_EXPORTS)
    Set expTab = CustomTable.Create(shExports.ListObjects(1), idCol:="Export Number")
    Set keysList = New BetterArray
    Set headersList = New BetterArray
    Set messageList = New BetterArray
    Set fileNameList = New BetterArray
    Set statusRng = expTab.DataRange("Status")

    Set check = Checking.Create("Export incoherences", "Type--Where?--Details")
    Set workFun = Application.WorksheetFunction
    Set vars = LLVariables.Create(this.dictionary)

    headersList.Push "Label button", "Password", "Export metadata sheets", _
                     "File format", "File name", "Header format"
    keysList.Push "exp-mis-lab", "exp-mis-pass", "exp-mis-meta", _
                  "exp-mis-form", "exp-mis-name", "exp-mis-head"
    messageList.Push EXP_MIS_LAB, EXP_MIS_PASS, EXP_MIS_META, _
                      EXP_MIS_FORM, EXP_MIS_NAME, EXP_MIS_HEAD

    checkingCounter = 0
    numberOfExports = statusRng.Rows.Count

    For counter = 1 To numberOfExports

        Set cellRng = expTab.CellRange("Status", counter + statusRng.Row - 1)
        expStatus = cellRng.Value
        pwd = expTab.Value(colName:="Password", keyName:=CStr(counter))
        exportKeepsId = expTab.Value(colName:="Include personal identifiers", keyName:=CStr(counter))

        For headerCounter = keysList.LowerBound To keysList.UpperBound
            If IsEmpty(expTab.CellRange(headersList.Item(headerCounter), cellRng.Row)) And (expStatus = "active") Then
                checkingCounter = checkingCounter + 1
                keyName = keysList.Item(headerCounter)
                infoMessage = FormatMessage(messageList.Item(headerCounter), cellRng.Row, counter)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
            End If
        Next headerCounter

        On Error Resume Next
            Set exportRng = Nothing
            Set exportRng = this.dictionary.DataRange("Export " & counter, includeHeaders:=False)
        On Error GoTo 0

        If (expStatus = "active") And (exportRng Is Nothing) Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-unfound-dictcolumn"
            infoMessage = FormatMessage(EXP_UNFOUND_DICTCOLUMN, cellRng.Row, counter)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        ElseIf (expStatus = "active") And (Not exportRng Is Nothing) Then
            If workFun.CountBlank(exportRng) = exportRng.Rows.Count Then
                checkingCounter = checkingCounter + 1
                keyName = "exp-act-empty"
                infoMessage = FormatMessage(EXP_ACT_EMPTY, cellRng.Row, counter)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (pwd <> "yes") And (exportKeepsId = "yes") Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-id-passwd"
            infoMessage = FormatMessage(EXP_ID_PASSWD, cellRng.Row, counter)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        End If

        actualFileName = expTab.Value(colName:="File Name", keyName:=counter)
        fileNameList.Clear
        fileNameList.Items = Split(actualFileName, "+")

        For fileCounter = fileNameList.LowerBound To fileNameList.UpperBound
            fileNameChunk = fileNameList.Item(fileCounter)

            If Not (InStr(1, fileNameChunk, Chr$(34)) > 0) Then
                If Not vars.Contains(fileNameChunk) Then
                    checkingCounter = checkingCounter + 1
                    keyName = "exp-filename-varunfound"
                    infoMessage = FormatMessage(EXP_FILENAME_VARUNFOUND, cellRng.Row, counter, fileNameChunk)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                ElseIf (vars.Value(colName:="Sheet Type", varName:=fileNameChunk) <> "vlist1D") Then
                    checkingCounter = checkingCounter + 1
                    keyName = "exp-filename-varnotvlist"
                    infoMessage = FormatMessage(EXP_FILENAME_VARNOTVLIST, cellRng.Row, counter, fileNameChunk)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                End If
            End If
        Next fileCounter
    Next counter

    this.checkTables.Push check
End Sub

'@sub-title Review translations table for missing labels.
Private Sub CheckTranslations()
    Dim translationsLo As ListObject
    Dim shTranslations As Worksheet
    Dim languageName As String
    Dim missingMessage As String
    Dim missingCount As Long
    Dim languageList As BetterArray
    Dim check As IChecking
    Dim counter As Long
    Dim columnRange As Range
    Dim transManager As ISetupTranslationsTable
    Dim shRegistry As Worksheet


    On Error GoTo TranslationFailed

    
    Set shTranslations = ResolveWorksheet(SHEET_TRANSLATIONS)
    Set shRegistry = ResolveWorksheet(SHEET_REGISTRY)
    Set translationsLo = shTranslations.ListObjects(1)
    Set transManager = SetupTranslationsTable.Create(translationsLo)
    Set check = Checking.Create("Translations", "Incoherence--Where?--Details")


    '
    this.passwordService.UnProtect SHEET_TRANSLATIONS
    transManager.SetDisplayPrompts False

    
    'Update the registry after passwords updates
    transManager.UpdateFromRegistry shRegistry
    

    ' If Not translationsLo.DataBodyRange Is Nothing Then
    '     For counter = 1 To headersRange.Columns.Count
    '         languageName = headersRange.Cells(1, counter).Value
    '         Set columnRange = translationsLo.ListColumns(languageName).DataBodyRange
    '         missingCount = Application.WorksheetFunction.CountBlank(columnRange)
    '         If missingCount > 0 Then
    '             missingMessage = "Translations Sheet--" & missingCount & _
    '                              " labels are missing for column " & languageName & "."
    '             check.Add "trads-mis-labs-" & counter, missingMessage, checkingInfo
    '         End If
    '     Next counter
    ' End If

    this.checkTables.Push check

TranslationsFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_TRANSLATIONS
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Run analysis table checks across all sections.
Private Sub CheckAnalysis()
    Dim analysisSheet As Worksheet
    Dim analysisManager As IAnalysis
    Dim selectionCell As Range
    Dim previousSelection As String
    Dim selectionConfigured As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set analysisSheet = ResolveWorksheet(SHEET_ANALYSIS)

    On Error Resume Next
        Set selectionCell = analysisSheet.Range("RNG_SelectTable")
        If Not selectionCell Is Nothing Then
            selectionConfigured = True
            previousSelection = CStr(selectionCell.Value)
            selectionCell.Value = "Add or remove rows of all tables"
        End If
    On Error GoTo 0

    On Error GoTo AnalysisFailed
        this.passwordService.UnProtect SHEET_ANALYSIS
        Set analysisManager = Analysis.Create(analysisSheet)
        analysisManager.RemoveRows
        this.passwordService.Protect SHEET_ANALYSIS
    On Error GoTo 0

    If selectionConfigured Then
        On Error Resume Next
            selectionCell.Value = previousSelection
        On Error GoTo 0
    End If

    ' CheckTable "Global summary"
    ' CheckTable "Univariate analysis"
    ' CheckTable "Bivariate analysis"
    ' CheckTable "Time series analysis"
    ' CheckTable "Time series graphs"
    ' CheckTable "Spatial analysis"
    ' CheckTable "Spatio-temporal analysis"
    Exit Sub

AnalysisFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_ANALYSIS
        If selectionConfigured Then selectionCell.Value = previousSelection
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Evaluate a specific analysis section for structural errors.
' Private Sub CheckTable(ByVal partName As String)

'     Const TAB_GLOBAL_SUMMARY As String = "Tab_global_summary"
'     Const TAB_UNIVARIATE As String = "Tab_Univariate_Analysis"
'     Const TAB_BIVARIATE As String = "Tab_Bivariate_Analysis"
'     Const TAB_TIME_SERIES As String = "Tab_TimeSeries_Analysis"
'     Const TAB_SPATIAL As String = "Tab_Spatial_Analysis"
'     Const TAB_SPATIOTEMP As String = "Tab_SpatioTemporal_Analysis"
'     Const TAB_TIME_SERIES_GRAPHS As String = "Tab_Graph_TimeSeries"

'     Dim listObject As ListObject
'     Dim listName As String
'     Dim headerRange As Range
'     Dim specsRange As Range
'     Dim specs As ITablesSpecs
'     Dim analysisSheet As Worksheet
'     Dim check As IChecking
'     Dim rowIndex As Long
'     Dim checkingCounter As Long
'     Dim keyName As String
'     Dim infoMessage As String
'     Dim analysisFormula As String
'     Dim workFun As WorksheetFunction
'     Dim tableScope As Byte

'     listName = Switch( _
'         partName = "Global summary", TAB_GLOBAL_SUMMARY, _
'         partName = "Univariate analysis", TAB_UNIVARIATE, _
'         partName = "Bivariate analysis", TAB_BIVARIATE, _
'         partName = "Time series analysis", TAB_TIME_SERIES, _
'         partName = "Spatial analysis", TAB_SPATIAL, _
'         partName = "Spatio-temporal analysis", TAB_SPATIOTEMP, _
'         partName = "Time series graphs", TAB_TIME_SERIES_GRAPHS _
'     )

'     Set analysisSheet = ResolveWorksheet(SHEET_ANALYSIS)
'     Set workFun = Application.WorksheetFunction
'     checkingCounter = 0

'     On Error Resume Next
'         Set listObject = analysisSheet.ListObjects(listName)
'     On Error GoTo 0
'     If listObject Is Nothing Then Exit Sub

'     Set check = Checking.Create("Analysis incoherences", _
'                                 partName & "--Where?--Details")

'     Set headerRange = listObject.HeaderRowRange

'     For rowIndex = 1 To listObject.ListRows.Count
'         Set specsRange = listObject.ListRows(rowIndex).Range

'         If ((workFun.CountBlank(specsRange) = specsRange.Columns.Count) And _
'             (listName <> TAB_TIME_SERIES)) Or _
'            (workFun.CountBlank(specsRange) >= specsRange.Columns.Count - 1) Then

'             keyName = "ana-empty-tab"
'             checkingCounter = checkingCounter + 1
'             infoMessage = FormatMessage(ANA_EMPTY_TAB, specsRange.Row)
'             check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
'         Else

'             Set specs = TablesSpecs.Create(headerRange, specsRange, this.dictionary, this.choices)
'             tableScope = specs.TableType

'             If (Not specs.ValidTable()) Then
'                 checkingCounter = checkingCounter + 1
'                 keyName = "ana-inv-tab"
'                 infoMessage = FormatMessage(ANA_INV_TAB, specsRange.Row, specs.ValidityReason())
'                 check.Add keyName & "-" & checkingCounter, infoMessage, checkingError

'                 If (tableScope = TypeTimeSeries) And (specs.IsNewSection()) _
'                    And (workFun.CountBlank(specsRange) < specsRange.Columns.Count - 1) Then
'                     keyName = "ana-ts-newsec"
'                     infoMessage = FormatMessage(ANA_TS_NEWSEC, specsRange.Row)
'                     check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
'                 End If
'             End If

'             If (specs.Value("title") = vbNullString) And _
'                (tableScope <> TypeGlobalSummary) And _
'                (tableScope <> TypeTimeSeriesGraph) Then

'                 checkingCounter = checkingCounter + 1
'                 keyName = "ana-empty-title"
'                 infoMessage = FormatMessage(ANA_EMPTY_TITLE, specsRange.Row)
'                 check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
'             End If

'             If (tableScope = TypeUnivariate) And _
'                specs.HasGraph() And _
'                (specs.Value("flip") = "yes") And _
'                (specs.Value("percentage") = "yes") Then

'                 checkingCounter = checkingCounter + 1
'                 keyName = "ana-uaflip-perc"
'                 infoMessage = FormatMessage(ANA_UAFLIP_PERC, specsRange.Row)
'                 check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
'             End If

'             If (tableScope <> TypeTimeSeries) And (tableScope <> TypeTimeSeriesGraph) Then
'                 If (Not specs.HasGraph()) And (specs.Value("flip") = "yes") Then
'                     checkingCounter = checkingCounter + 1
'                     keyName = "ana-adgr-flip"
'                     infoMessage = FormatMessage(ANA_ADGR_FLIP, specsRange.Row)
'                     check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
'                 End If
'             End If

'             If (tableScope <> TypeTimeSeriesGraph) Then
'                 analysisFormula = specs.Value("function")
'                 If (analysisFormula <> vbNullString) Then
'                     keyName = "ana-incor-form"
'                     infoMessage = FormulaMessage(analysisFormula, ANA_INCOR_FORM, specsRange.Row, _
'                                                  formulaType:="analysis")
'                     If (infoMessage <> vbNullString) Then
'                         checkingCounter = checkingCounter + 1
'                         check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
'                     End If
'                 End If
'             End If
'         End If
'     Next rowIndex

'     this.checkTables.Push check
' End Sub

'@sub-title Publish all collected checkings to the reporting worksheet.
Private Sub PrintReport()
    Dim checkOutput As ICheckingOutput
    Dim reportSheet As Worksheet
    Dim dropdowns As IDropdownLists
    Dim formatRange As Range

    Set reportSheet = ResolveWorksheet(SHEET_CHECK_OUTPUT)
    reportSheet.Cells.Clear
    Set checkOutput = CheckingOutput.Create(reportSheet)
    Set dropdowns = DropdownLists.Create(ResolveWorksheet(SHEET_DROPDOWN))

    reportSheet.Cells.EntireRow.Hidden = False

    checkOutput.PrintOutput this.checkTables
    
    On Error Resume Next
    reportSheet.Activate
End Sub

'@section Message helpers
'===============================================================================
Private Function FormatMessage(ByVal messageTag As String, _
                               Optional ByVal valueOne As String = vbNullString, _
                               Optional ByVal valueTwo As String = vbNullString, _
                               Optional ByVal valueThree As String = vbNullString, _
                               Optional ByVal valueFour As String = vbNullString) As String

    Dim template As String

    EnsureInitialised

    template = ReplacePlaceholder(messageTag, "%1", valueOne)
    template = ReplacePlaceholder(template, "%2", valueTwo)
    template = ReplacePlaceholder(template, "%3", valueThree)
    template = ReplacePlaceholder(template, "%4", valueFour)

    FormatMessage = template
End Function

Private Function FormulaMessage(ByVal formValue As String, _
                                ByVal messageTag As String, _
                                Optional ByVal valueOne As String = vbNullString, _
                                Optional ByVal valueTwo As String = vbNullString, _
                                Optional ByVal formulaType As String = "linelist") As String
    Dim setupFormulas As IFormulas

    If LenB(Trim$(formValue)) = 0 Then Exit Function

    On Error GoTo FormulaFail
        Set setupFormulas = Formulas.Create(this.dictionary, this.formulaStore, formValue)
        If Not setupFormulas.Valid(formulaType:=formulaType) Then
            FormulaMessage = FormatMessage(messageTag, valueOne, valueTwo, setupFormulas.reason())
        End If
    On Error GoTo 0
    Exit Function

FormulaFail:
    ' If validation fails unexpectedly, surface the failure as a formatted message.
    FormulaMessage = FormatMessage(messageTag, valueOne, valueTwo, Err.Description)
    Err.Clear
End Function

Private Function ReplacePlaceholder(ByVal template As String, _
                                    ByVal placeholder As String, _
                                    ByVal replacement As String) As String
    If LenB(template) = 0 Then
        ReplacePlaceholder = vbNullString
    ElseIf InStr(1, template, placeholder, vbBinaryCompare) = 0 Then
        ReplacePlaceholder = template
    Else
        ReplacePlaceholder = Replace(template, placeholder, replacement)
    End If
End Function

'@section Interface implementation
'===============================================================================
Private Sub ISetupErrors_Initialise(ByVal hostBook As Workbook, _
                                    Optional ByVal passwordService As IPasswords, _
                                    Optional ByVal dictionary As ILLdictionary, _
                                    Optional ByVal choices As ILLChoices, _
                                    Optional ByVal formulaStore As IFormulaData, _
                                    Optional ByVal stateScope As IApplicationState)
    Initialise hostBook, passwordService, dictionary, choices, formulaStore, stateScope
End Sub

Private Sub ISetupErrors_Run()
    Run
End Sub

Private Property Get ISetupErrors_Checkings() As BetterArray
    Set ISetupErrors_Checkings = Checkings
End Property

'@label:ThrowError
'@description Raise a ProjectError aligned error from the current class.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASS_NAME, "Error: " & errorMessage
End Sub