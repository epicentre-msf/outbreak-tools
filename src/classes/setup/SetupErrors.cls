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
'@depends Analysis, ApplicationState, BetterArray, Checking, CheckingOutput, ChoiceFormula, CustomTable, FormulaData, Formulas, LLChoices, LLVariables, LLdictionary, Passwords, SetupTranslationsTable, TableSpecs

Option Explicit

'@class SetupErrors
'@description
'Checks the setup sheets and publishes a report of what it finds. Create binds
'the workbook, Initialise sets the run up, Run walks every check, and Checkings
'answers the entries. The dictionary is checked for duplicate variables,
'missing labels, unknown controls and formulas that fail to parse; the choices
'for unused entries and missing labels; the exports for missing metadata, for
'an "Admin levels" cell the export cannot read, and for variables the dictionary
'has none of; and the translations for missing and duplicate labels. Entries are
'collected as Checking objects and printed to the __checkRep worksheet through
'CheckingOutput.
'
'THE CHECKS FOLLOW THE BUILDER
'-------------------------------------------------------------------------------
'Every check here answers what the linelist build will do with the same setup
'sheet. Three rules are copied from the class that builds them: a file-name
'template is parsed the way LLExport.ResolveTemplateChunk parses it, a formula
'is explained in the context it was validated in, and a choice name is matched
'without regard to case, the way LLChoices.Categories matches it at build time.
'
'THE SHEETS ARE TIDIED FIRST
'-------------------------------------------------------------------------------
'Each checker sorts and resizes the sheet it is about to read, so every check
'below reads a sheet already in shape. The column arrays CheckDictionary works
'from are therefore read AFTER the sort and the resize: the tidy-up moves rows,
'and an array taken before it points at rows that have moved.
'
'THE REPORT SHEET IS WRITABLE
'-------------------------------------------------------------------------------
'__checkRep sits outside the hidden list and outside the protection list, and
'Passwords returns early for it by name. PrintReport clears it and writes to it
'directly, which is what the setup ribbon's check button reaches.

'@section State
'===============================================================================
Private Type TSetupErrorsState
    hostBook As Workbook              'Target workbook containing setup data
    passwordService As Passwords     'Password handler protecting/unprotecting sheets
    dictionary As LLdictionary       'Dictionary accessor for setup data
    choices As LLChoices             'Choices accessor for list validation
    formulaStore As FormulaData      'Formula data repository used for validation
    checkTables As BetterArray        'Collection of generated checkings
    appScope As ApplicationState     'Application state guard ensuring UI stays responsive
    transManager As SetupTranslationsTable 'translation table accessor for setup data
    formulaChecker As Formulas       'One parser reused for every formula of a run
    ownsBusyState As Boolean          'True when this object is the one that entered the busy state
    sealed As Boolean                 'Set by the factory, closes Initialise
End Type

Private this As TSetupErrorsState

Private Const CLASS_NAME As String = "SetupErrors"

Private Const SHEET_FORMULA As String = "__formula"
Private Const SHEET_PASSWORDS As String = "__pass"
Private Const SHEET_CHECK_OUTPUT As String = "__checkRep"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const SHEET_REGISTRY As String = "__updated"

'Separator wrapping every entry of a delimited lookup index. It sits on both
'sides of each name, so a short name can never match inside a longer one.
Private Const INDEX_SEPARATOR As String = vbNullChar

Private Const CONTEXT_LINELIST As String = "linelist"
Private Const CONTEXT_ANALYSIS As String = "analysis"

'@section List of errors to take in account
'===============================================================================

Private Const ANA_ADGR_FLIP As String               = "Analysis, line %1--There is no graph for this table, so ""flip coordinates"" will be ignored."
Private Const ANA_EMPTY_TAB As String               = "Analysis, line %1--This line is completely empty, the table will be ignored in analysis."
Private Const ANA_EMPTY_TITLE As String             = "Analysis, line %1--The table does not have a title."
Private Const ANA_INCOR_FORM As String              = "Analysis, line %1--The formula specified for the table will eventually fail and not be parsed. %3"
Private Const ANA_INV_TAB As String                 = "Analysis, line %1--The table is invalid. There is an error in specifications: %2"
Private Const ANA_SPECS_FAILED As String            = "Analysis, line %1--The specifications of this table could not be read, so the line was skipped: %2"
Private Const ANA_TIDY_FAILED As String             = "Analysis, whole sheet--The empty analysis rows could not be removed before the checks: %1"
Private Const ANA_TS_NEWSEC As String               = "Analysis, line %1--This table is invalid and is a new section of time series. This will likely generate errors or impact all the time series sheet."
Private Const ANA_UAFLIP_PERC As String             = "Analysis, line %1--The graph for this univariate table will not contain percentages because ""flip coordinates"" = yes."
Private Const CHOI_EMPTYCHOI_LAB As String          = "Choices, line %1--There is no attached ""List Name"" value for Label ""%2""."
Private Const CHOI_EMPTYCHOI_ORDER As String        = "Choices, line %1--There is no attached ""List Name"" value for Ordering list ""%2""."
Private Const CHOI_EMPTY_ORDER As String            = "Choices, line %1--There is no value of ""Ordering list"" for choice ""%2""."
Private Const CHOI_MIS_COLUMN As String             = "Choices, header row--The column ""%1"" is missing, so the checks that need it were skipped."
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
Private Const DICT_MIS_COLUMN As String             = "Dictionary, header row--The column ""%1"" is missing, so the checks that need it were skipped."
Private Const DICT_UNIQUE_VAL As String             = "Dictionary, line %1--Duplicate values of variable %2 will be highlighted in red because Unique = ""yes""."
Private Const DICT_UNKNOWN_CONTROL As String        = "Dictionary, line %1--The control ""%2"" of variable ""%3"" is unknown."
Private Const DICT_VALID_CONTROL As String          = "Dictionary, line %1--Validation requires the type of the variable, please consider adding the type for variable %2."
Private Const DICT_VAR_LENGTH As String             = "Dictionary, line %1--Variable ""%2"" length is less than 4. Variable should have length >= 4."
Private Const DICT_VAR_UNIQUE As String             = "Dictionary, line %1--Variable ""%2"" is duplicate, variable names should be unique."
Private Const EXP_ACT_EMPTY As String               = "Exports, line %1--The export is active, but the corresponding column is empty in the dictionary. You may end up with empty data in this export. Please correct the status of export %2."
Private Const EXP_ADM_LEVELS As String              = "Exports, line %1--The Export Number %2 has an ""Admin levels"" value that cannot be read: ""%3"". Write the levels joined with a plus, like ""admin1+admin2"", or leave the cell empty to export every admin level. Please correct it."
Private Const EXP_FILENAME_VARNOTVLIST As String    = "Exports, line %1--The Export Number %2 file name contains a variable named ""%3"" which is not a vlist1D variable, but a hlist2D. It will be considered as text in exports."
Private Const EXP_FILENAME_VARUNFOUND As String     = "Exports, line %1--The Export Number %2 file name contains a variable named ""%3"" which does not exists. It will be considered as text in exports."
Private Const EXP_ID_PASSWD As String               = "Exports, line %1--The Export Number %2 has no password, but keeps identifiers."
Private Const EXP_MIS_COLUMN As String              = "Exports, header row--The column ""%1"" is missing, so the checks that need it were skipped."
Private Const EXP_MIS_FORM As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""File Format"" (xlsx/xlsb) attached to it."
Private Const EXP_MIS_HEAD As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Header format"" attached to it. It will default to ""variables names""."
Private Const EXP_MIS_LAB As String                 = "Exports, line %1--The Export Number %2 is active, but there is no label attached to it."
Private Const EXP_MIS_META As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Export Metadata"" (yes/no) attached to it."
Private Const EXP_MIS_NAME As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""File name"" attached to it."
Private Const EXP_MIS_PASS As String                = "Exports, line %1--The Export Number %2 is active, but there is no value for ""Password"" (yes/no) attached to it."
Private Const EXP_MIS_STATUS As String              = "Exports, header row--The ""Status"" column is missing, so no export could be checked."
Private Const EXP_UNFOUND_DICTCOLUMN As String      = "Exports, line %1--The Export Number %2 is active, but the corresponding column in the dictionary is missing."


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and initialise a setup checker instance
'@details
'Factory method on the predeclared instance. Creates a new SetupErrors,
'calls Initialise with the supplied workbook and optional collaborators,
'seals it, then returns the newly built instance.
'@param hostBook Workbook. The workbook containing setup data.
'@param passwordService Optional Passwords. Sheet protection handler.
'@param dictionary Optional LLdictionary. Dictionary data accessor.
'@param choices Optional LLChoices. Choices data accessor.
'@param formulaStore Optional FormulaData. Formula validation handler.
'@param transManager Optional SetupTranslationsTable. Translations table accessor.
'@param stateScope Optional ApplicationState. Application state guard.
'@return SetupErrors. A fully initialised instance ready to run checks.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Function Create(ByVal hostBook As Workbook, _
                       Optional ByVal passwordService As Passwords, _
                       Optional ByVal dictionary As LLdictionary, _
                       Optional ByVal choices As LLChoices, _
                       Optional ByVal formulaStore As FormulaData, _
                       Optional ByVal transManager As SetupTranslationsTable, _
                       Optional ByVal stateScope As ApplicationState) As SetupErrors
    Dim instance As SetupErrors

    Set instance = New SetupErrors
    instance.Initialise hostBook, passwordService, dictionary, choices, formulaStore, transManager, stateScope
    instance.Seal

    Set Create = instance
End Function

'@label:seal
'@sub-title Close the instance against a second Initialise
'@details
'Called by the factory. A holder that could call Initialise again could swap
'the workbook between Run and Checkings and read results that belong to
'another file. LLdictionary, LLChoices, Passwords, TableSpecs, Checking,
'CustomTable, ChoiceFormula and DropdownLists all carry the same guard.
'@export
Public Sub Seal()
    this.sealed = True
End Sub

'@section Initialisation
'===============================================================================

'@label:initialise
'@sub-title Bind the checker to a workbook and optional collaborators
'@details
'Validates the workbook reference, stores it along with optional
'collaborators, then calls ConfigureDependencies to create any missing
'default instances and ResetCheckContainer to prepare the results array.
'@param hostBook Workbook. The workbook containing setup tables.
'@param passwordService Optional Passwords. Sheet protection handler.
'@param dictionary Optional LLdictionary. Dictionary data accessor.
'@param choices Optional LLChoices. Choices data accessor.
'@param formulaStore Optional FormulaData. Formula validation handler.
'@param transManager Optional SetupTranslationsTable. Translations table accessor.
'@param stateScope Optional ApplicationState. Application state guard.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
'@export
Public Sub Initialise(ByVal hostBook As Workbook, _
                      Optional ByVal passwordService As Passwords, _
                      Optional ByVal dictionary As LLdictionary, _
                      Optional ByVal choices As LLChoices, _
                      Optional ByVal formulaStore As FormulaData, _
                      Optional ByVal transManager As SetupTranslationsTable, _
                      Optional ByVal stateScope As ApplicationState)

    GuardNotSealed "Initialise"
    ValidateWorkbook hostBook
    Set this.hostBook = hostBook

    Set this.passwordService = passwordService
    Set this.dictionary = dictionary
    Set this.choices = choices
    Set this.formulaStore = formulaStore
    Set this.transManager = transManager
    Set this.appScope = stateScope

    ConfigureDependencies
    ResetCheckContainer
End Sub

'@section Execution
'===============================================================================

'@label:run
'@sub-title Execute all setup checks and publish results
'@details
'Suspends Excel UI via EnterBusyState, resets the check container, then
'runs each checker in sequence: CheckDictionary, CheckChoice, CheckExports,
'CheckAnalysis, CheckTranslations. On completion, publishes all collected
'Checking objects to the reporting worksheet via PrintReport. Restores Excel
'state even when an error occurs during the checking run.
'The handler is armed above EnterBusyState. ApplicationState.ApplyBusyState
'switches ScreenUpdating and DisplayAlerts off as its first two statements and
'raises afterwards if a later assignment fails, so a raise in there has to
'reach LeaveBusyState. Otherwise the user is left in an Excel session that no
'longer repaints.
'@export
Public Sub Run()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    EnsureInitialised

    On Error GoTo RunFailed
        EnterBusyState

        ResetCheckContainer
        Set this.formulaChecker = Nothing

        CheckDictionary
        CheckChoice
        CheckExports
        CheckAnalysis
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

'@label:checkings
'@prop-title Collection of generated checking objects
'@details
'Returns a clone of the BetterArray holding the Checking instances produced
'by the most recent Run call, so a caller that clears it leaves this object's
'own results alone. Formulas.Varlists and LLChoices.Categories answer the
'same way. Raises when the checker has not been initialised.
'@return BetterArray. A copy of the generated Checking instances.
'@export
Public Property Get Checkings() As BetterArray
    Dim copy As BetterArray

    EnsureInitialised

    'Clone reads the internal array, which is unallocated while nothing has
    'been pushed, so an empty container answers with a fresh empty array.
    If this.checkTables.Length = 0 Then
        Set copy = New BetterArray
    Else
        Set copy = this.checkTables.Clone()
    End If

    Set Checkings = copy
End Property

'@section Helper workflow
'===============================================================================

'@label:ensure-initialised
'@sub-title Ensure the checker has been initialised
'@throws ProjectError.ObjectNotInitialized When the workbook reference is Nothing.
Private Sub EnsureInitialised()
    If (this.hostBook Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupErrors.Initialise must be called before usage."
    End If
End Sub

'@label:guard-not-sealed
'@sub-title Reject a second Initialise on a sealed instance
'@param propName String. Name of the routine being guarded.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then
        ThrowError ProjectError.SomethingWentWrong, _
                   propName & " can not run on a sealed SetupErrors instance."
    End If
End Sub

'@label:configure-dependencies
'@sub-title Prepare collaborators when not provided
'@details
'Creates default instances for any optional collaborator that was not
'supplied during Initialise: Passwords, LLdictionary, LLChoices, FormulaData
'and SetupTranslationsTable. The __formula worksheet is resolved once, and
'only when the formula store has to be built.
Private Sub ConfigureDependencies()

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
        Set this.formulaStore = FormulaData.Create(ResolveWorksheet(SHEET_FORMULA))
    End If

    If this.transManager Is Nothing Then
        Set this.transManager = SetupTranslationsTable.Create(ResolveWorksheet(SHEET_TRANSLATIONS).ListObjects(1))
    End If
End Sub

'@label:reset-check-container
'@sub-title Ensure the container for checkings exists and is empty
Private Sub ResetCheckContainer()
    If this.checkTables Is Nothing Then
        Set this.checkTables = New BetterArray
    Else
        this.checkTables.Clear
    End If
End Sub

'@label:application-scope
'@sub-title Retrieve or create the application state scope
'@return ApplicationState. The cached or newly created scope.
Private Function ApplicationScope() As ApplicationState
    If this.appScope Is Nothing Then
        Set this.appScope = ApplicationState.Create(Application)
    End If

    Set ApplicationScope = this.appScope
End Function

'@label:enter-busy-state
'@sub-title Suspend Excel UI features while running checks
'@details
'A scope that is already busy belongs to the caller who made it busy, so this
'object leaves it alone and LeaveBusyState will not restore it. The ownership
'flag is written BEFORE ApplyBusyState, because a failure inside that call
'still leaves settings to put back.
Private Sub EnterBusyState()
    Dim scope As ApplicationState

    Set scope = ApplicationScope()
    If scope.IsBusy Then Exit Sub

    this.ownsBusyState = True
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False
End Sub

'@label:leave-busy-state
'@sub-title Restore Excel UI configuration
'@details
'Restores only when this object was the one that entered the busy state.
'ApplicationState.Restore puts the whole snapshot back and clears IsBusy every
'time it runs, so restoring a scope handed in by a caller would drop that
'caller's own busy block in the middle of its work.
'@param silent Optional Boolean. When True, suppresses errors during restore. Defaults to False.
Private Sub LeaveBusyState(Optional ByVal silent As Boolean = False)
    Dim scope As ApplicationState

    If Not this.ownsBusyState Then Exit Sub

    Set scope = this.appScope
    If scope Is Nothing Then Exit Sub

    this.ownsBusyState = False

    'ApplicationState already knows how to fail quietly, so hand the flag to it
    'instead of trapping the error again here.
    scope.Restore silent:=silent
End Sub

'@label:resolve-worksheet
'@sub-title Resolve a workbook worksheet by name
'@param sheetName String. The worksheet name to look up.
'@return Worksheet. The resolved worksheet.
'@throws ProjectError.ElementNotFound When the worksheet does not exist.
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

'@label:validate-workbook
'@sub-title Validate the incoming workbook reference
'@param hostBook Workbook. The workbook to validate.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
Private Sub ValidateWorkbook(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupErrors requires a workbook reference."
    End If
End Sub

'@section Column readers
'===============================================================================

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'A cell holding an error value raises on CStr, and so does a Null. Every read
'in this class goes through here.
'@param cellValue Variant. The raw cell value.
'@return String. The text, or an empty string when there is none.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:column-text
'@sub-title Read a one-column Range into a 1-based array of text
'@details
'A missing column answers an array of empty strings of the requested length,
'so the loop that reads it can index it freely. One worksheet read replaces
'one read per row.
'@param columnRng Range. The column to read, or Nothing.
'@param rowCount Long. Number of rows the caller walks.
'@return Variant. A String array indexed 1 to rowCount.
Private Function ColumnText(ByVal columnRng As Range, ByVal rowCount As Long) As Variant
    Dim values() As String
    Dim rawValues As Variant
    Dim counter As Long

    If rowCount < 1 Then rowCount = 1
    ReDim values(1 To rowCount)

    If columnRng Is Nothing Then
        ColumnText = values
        Exit Function
    End If

    'A one-cell Range answers a scalar rather than a two-dimensional array.
    If columnRng.Cells.Count = 1 Then
        values(1) = CellText(columnRng.Value)
        ColumnText = values
        Exit Function
    End If

    rawValues = columnRng.Value

    For counter = 1 To rowCount
        If counter <= UBound(rawValues, 1) Then
            values(counter) = CellText(rawValues(counter, 1))
        End If
    Next

    ColumnText = values
End Function

'@label:read-table-column
'@sub-title Read one named column of a table into an array
'@details
'The lookup is strict, so a request for "Control" can never answer with
'"Control Details" and a request for "Export 1" can never answer with
'"Export 10".
'@param tabl CustomTable. The table to read.
'@param colName String. The exact column header.
'@param rowCount Long. Number of rows the caller walks.
'@param found Boolean. Set to True when the column exists.
'@return Variant. A String array indexed 1 to rowCount.
Private Function ReadTableColumn(ByVal tabl As CustomTable, _
                                 ByVal colName As String, _
                                 ByVal rowCount As Long, _
                                 ByRef found As Boolean) As Variant
    Dim columnRng As Range

    Set columnRng = tabl.DataRange(colName, strictSearch:=True)
    found = Not (columnRng Is Nothing)
    ReadTableColumn = ColumnText(columnRng, rowCount)
End Function

'@label:build-index
'@sub-title Build a delimited lookup index from a list of names
'@details
'The separator sits on both sides of every entry, so InStr on
'separator & name & separator can never match part of a longer name. This
'replaces a linear scan that copied a whole BetterArray on each call.
'@param values BetterArray. The names to index.
'@return String. The delimited index.
Private Function BuildIndex(ByVal values As BetterArray) As String
    Dim builder As String
    Dim counter As Long

    builder = INDEX_SEPARATOR
    If values Is Nothing Then
        BuildIndex = builder
        Exit Function
    End If
    If values.Length = 0 Then
        BuildIndex = builder
        Exit Function
    End If

    For counter = values.LowerBound To values.UpperBound
        builder = builder & CStr(values.Item(counter)) & INDEX_SEPARATOR
    Next

    BuildIndex = builder
End Function

'@label:index-holds
'@sub-title Test whether a delimited index holds one name
'@details
'The comparison ignores case, because LLChoices.Categories reads the Choices
'sheet with AutoFilter criteria and those ignore case. The build therefore
'finds a list whose name differs only in case, and the checker follows it.
'@param indexValue String. A delimited index built by BuildIndex.
'@param name String. The name to look for.
'@return Boolean. True when the name is in the index.
Private Function IndexHolds(ByVal indexValue As String, ByVal nameValue As String) As Boolean
    If LenB(nameValue) = 0 Then Exit Function
    IndexHolds = (InStr(1, indexValue, INDEX_SEPARATOR & nameValue & INDEX_SEPARATOR, vbTextCompare) > 0)
End Function

'@label:count-names
'@sub-title Count how many times each name appears
'@details
'One pass over the array replaces one WorksheetFunction.CountIf per row, each
'of which scanned the whole column inside Excel. A Collection matches its keys
'without regard to case, which is how CountIf compared too.
'@param values Variant. A 1-based array of names.
'@param rowCount Long. Number of entries to read.
'@return Collection. Counts keyed by name.
Private Function CountNames(ByVal values As Variant, ByVal rowCount As Long) As Collection
    Dim counts As Collection
    Dim counter As Long
    Dim keyName As String
    Dim seen As Long

    Set counts = New Collection

    For counter = 1 To rowCount
        keyName = CStr(values(counter))
        If LenB(keyName) > 0 Then
            seen = 0
            On Error Resume Next
                seen = counts.Item("k" & keyName)
            On Error GoTo 0
            If seen > 0 Then counts.Remove "k" & keyName
            counts.Add seen + 1, "k" & keyName
        End If
    Next

    Set CountNames = counts
End Function

'@label:name-count
'@sub-title Read one count out of the counts built by CountNames
'@param counts Collection. Counts keyed by name.
'@param name String. The name to read.
'@return Long. The number of rows holding that name.
Private Function NameCount(ByVal counts As Collection, ByVal nameValue As String) As Long
    If LenB(nameValue) = 0 Then Exit Function
    On Error Resume Next
        NameCount = counts.Item("k" & nameValue)
    On Error GoTo 0
End Function

'@label:export-identifier-number
'@sub-title Read the number out of one "Export Number" cell
'@details
'`LLExport` writes that cell through `ExportColumnName`, so it holds "Export 1"
'rather than "1", and `LLExport.ParseExportIdentifier` reads it back the same
'way. A bare number is taken as well, because that is what a hand-typed row
'holds.
'@param rawValue Variant. The value of the cell.
'@return Long. The export number, or 0 when the text carries none.
Private Function ExportIdentifierNumber(ByVal rawValue As Variant) As Long
    Dim text As String

    text = LCase$(Trim$(CellText(rawValue)))
    If LenB(text) = 0 Then Exit Function

    If Left$(text, 6) = "export" Then text = Trim$(Mid$(text, 7))
    If LenB(text) = 0 Then Exit Function
    If Not IsNumeric(text) Then Exit Function

    If CDbl(text) > 0 Then ExportIdentifierNumber = CLng(CDbl(text))
End Function

'@label:export-statuses
'@sub-title Status of every export row, keyed by the number the row carries
'@details
'Built once so the dictionary loop can ask about an export number without
'searching the Exports table again for each one. A number appearing twice
'keeps the status of its first row.
'@param expTab CustomTable. The Exports table.
'@return Collection. Statuses keyed by "k" & the export number.
Private Function ExportStatuses(ByVal expTab As CustomTable) As Collection
    Dim statuses As Collection
    Dim numberRng As Range
    Dim statusRng As Range
    Dim rowCount As Long
    Dim counter As Long
    Dim exportNumber As Long

    Set statuses = New Collection
    Set ExportStatuses = statuses

    If expTab Is Nothing Then Exit Function

    Set numberRng = expTab.DataRange("Export Number", strictSearch:=True)
    Set statusRng = expTab.DataRange("Status", strictSearch:=True)
    If (numberRng Is Nothing) Or (statusRng Is Nothing) Then Exit Function

    rowCount = numberRng.Rows.Count
    If statusRng.Rows.Count < rowCount Then rowCount = statusRng.Rows.Count

    For counter = 1 To rowCount
        exportNumber = ExportIdentifierNumber(numberRng.Cells(counter, 1).Value)
        If exportNumber > 0 Then
            On Error Resume Next
                statuses.Add CellText(statusRng.Cells(counter, 1).Value), "k" & exportNumber
            On Error GoTo 0
        End If
    Next counter
End Function

'@label:export-status-of
'@sub-title Read the status of one export number
'@param statuses Collection. The statuses built by ExportStatuses.
'@param exportNumber Long. The export number to read.
'@param statusValue String. Receives the status, empty when there is no row.
'@return Boolean. True when the Exports table holds a row for that number.
Private Function ExportStatusOf(ByVal statuses As Collection, _
                                ByVal exportNumber As Long, _
                                ByRef statusValue As String) As Boolean
    statusValue = vbNullString
    If statuses Is Nothing Then Exit Function

    On Error Resume Next
        Err.Clear
        statusValue = statuses.Item("k" & exportNumber)
        ExportStatusOf = (Err.Number = 0)
    On Error GoTo 0
End Function

'@label:known-control
'@sub-title Test whether a control name is one this build understands
'@param controlValue String. The control read from the dictionary.
'@return Boolean. True for one of the ten known controls.
Private Function KnownControl(ByVal controlValue As String) As Boolean
    Select Case controlValue
    Case "choice_manual", "choice_formula", "formula", "geo", "hf", _
         "custom", "list_auto", "case_when", "choice_custom", "choice_multiple"
        KnownControl = True
    End Select
End Function

'@label:contains-ignoring-case
'@sub-title Test membership of a BetterArray without regard to case
'@details
'BetterArray.Includes ends in a plain equality under binary compare, so a
'category written Yes in a CHOICE_FORMULA against yes on the Choices sheet
'reads as absent. The build matches the two, so the checker does too.
'@param values BetterArray. The values to search.
'@param target String. The value to look for.
'@return Boolean. True when the value is present in any case.
Private Function ContainsIgnoringCase(ByVal values As BetterArray, ByVal target As String) As Boolean
    Dim counter As Long

    If values Is Nothing Then Exit Function
    If values.Length = 0 Then Exit Function

    For counter = values.LowerBound To values.UpperBound
        If StrComp(CStr(values.Item(counter)), target, vbTextCompare) = 0 Then
            ContainsIgnoringCase = True
            Exit Function
        End If
    Next
End Function

'@label:is-admin-levels-cell
'@sub-title Test whether an "Admin levels" cell is written the way an export reads it
'@details
'Mirrors LLExporter.AdminLevelOfChunk: the cell holds the admin levels the
'export writes, joined with a plus and written `admin1+admin2`. Each chunk is
'trimmed and read without regard to case, so `Admin2` and ` admin2` both pass,
'and the level runs from 1 to 4 because LLdictionary.AppendGeoLines writes four.
'An empty cell passes: the export reads it as every level.
'The exporter honours the cell only when the whole of it reads, so a cell this
'reports is a cell the export ignores, and every geo column goes out of it.
'@param levelsText String. What the "Admin levels" cell of one export row holds.
'@return Boolean. True when the export can honour the cell.
Private Function IsAdminLevelsCell(ByVal levelsText As String) As Boolean

    Const LEVEL_PREFIX As String = "admin"
    Const LEVEL_SEPARATOR As String = "+"
    Const LEVEL_COUNT As Long = 4

    Dim chunks As Variant
    Dim chunkText As String
    Dim tail As String
    Dim counter As Long

    If LenB(Trim$(levelsText)) = 0 Then
        IsAdminLevelsCell = True
        Exit Function
    End If

    chunks = Split(levelsText, LEVEL_SEPARATOR)

    For counter = LBound(chunks) To UBound(chunks)
        chunkText = Trim$(CStr(chunks(counter)))
        If Len(chunkText) <> Len(LEVEL_PREFIX) + 1 Then Exit Function
        If StrComp(Left$(chunkText, Len(LEVEL_PREFIX)), _
                   LEVEL_PREFIX, vbTextCompare) <> 0 Then Exit Function
        tail = Right$(chunkText, 1)
        If tail < "1" Or tail > CStr(LEVEL_COUNT) Then Exit Function
    Next counter

    IsAdminLevelsCell = True
End Function

'@label:is-literal-chunk
'@sub-title Test whether a file-name chunk is a quoted literal
'@details
'Mirrors LLExport.IsLiteralTemplateChunk: the chunk is a literal when it
'starts AND ends with the same quote character. Testing for a quote anywhere
'in the chunk reported a single-quoted label as a missing variable.
'@param chunk String. One trimmed piece of the file-name template.
'@return Boolean. True when the chunk is a literal.
Private Function IsLiteralChunk(ByVal chunk As String) As Boolean
    Dim firstChar As String
    Dim lastChar As String

    If Len(chunk) < 2 Then Exit Function

    firstChar = Left$(chunk, 1)
    lastChar = Right$(chunk, 1)

    If (firstChar = Chr$(34)) And (lastChar = Chr$(34)) Then
        IsLiteralChunk = True
        Exit Function
    End If

    If (firstChar = "'") And (lastChar = "'") Then
        IsLiteralChunk = True
    End If
End Function

'@section Core checking routines
'===============================================================================

'@label:check-dictionary
'@sub-title Validate dictionary sheet for structure, duplicates, and formulas
'@details
'Unprotects the dictionary and exports sheets, tidies the dictionary table,
'then reads its ten columns once into arrays and walks the rows bottom-up
'checking for: duplicate variable names, short variable names, missing main
'labels, missing sheet names, unknown controls, choice formula mismatches,
'formula validity, min/max validation types, format/type consistency, and
'export column alignment. Collects all findings into a single Checking object
'and re-protects sheets on exit.
'The Choices sheet is only read here, through the choices index and the
'category lists, so it is left protected.
Private Sub CheckDictionary()

    Dim check As Checking
    Dim csTab As CustomTable
    Dim expTab As CustomTable
    Dim varRng As Range
    Dim expRng As Range
    Dim statusByNumber As Collection
    Dim workFun As WorksheetFunction
    Dim shDictionary As Worksheet
    Dim shExports As Worksheet
    Dim infoMessage As String
    Dim keyName As String
    Dim setupForm As ChoiceFormula
    Dim checkingCounter As Long
    Dim choiCategories As BetterArray
    Dim formCategories As BetterArray
    Dim choicesIndex As String
    Dim choiName As String
    Dim tabCounter As Long
    Dim catValue As String
    Dim expCounter As Long
    Dim firstRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim worksheetRow As Long
    Dim varValues As Variant
    Dim mainLabels As Variant
    Dim sheetNames As Variant
    Dim controlValues As Variant
    Dim controlDetails As Variant
    Dim minValues As Variant
    Dim maxValues As Variant
    Dim typeValues As Variant
    Dim formatValues As Variant
    Dim uniqueValues As Variant
    Dim nameCounts As Collection
    Dim varValue As String
    Dim controlValue As String
    Dim controlDetailsValue As String
    Dim minValue As String
    Dim maxValue As String
    Dim typeValue As String
    Dim exportStatus As String
    Dim hasMainLabel As Boolean
    Dim hasSheetName As Boolean
    Dim hasControl As Boolean
    Dim columnFound As Boolean

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

    this.passwordService.UnProtect SHEET_DICTIONARY
    this.passwordService.UnProtect SHEET_EXPORTS

    'The tidy-up comes first on purpose and the column arrays are read after
    'it: it moves rows, and arrays taken before it would point at rows that no
    'longer exist.
    csTab.RemoveRows includeIds:=False
    csTab.Sort "Sheet Name"

    Set varRng = csTab.DataRange("Variable Name", strictSearch:=True)

    If varRng Is Nothing Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Variable Name")
        check.Add "dict-mis-column-" & checkingCounter, infoMessage, checkingError
        this.checkTables.Push check
        this.passwordService.Protect SHEET_DICTIONARY
        this.passwordService.Protect SHEET_EXPORTS
        On Error GoTo 0
        Exit Sub
    End If

    firstRow = varRng.Row
    rowCount = varRng.Rows.Count

    varValues = ColumnText(varRng, rowCount)
    mainLabels = ReadTableColumn(csTab, "Main Label", rowCount, hasMainLabel)
    sheetNames = ReadTableColumn(csTab, "Sheet Name", rowCount, hasSheetName)
    controlValues = ReadTableColumn(csTab, "Control", rowCount, hasControl)
    controlDetails = ReadTableColumn(csTab, "Control Details", rowCount, columnFound)
    minValues = ReadTableColumn(csTab, "Min", rowCount, columnFound)
    maxValues = ReadTableColumn(csTab, "Max", rowCount, columnFound)
    typeValues = ReadTableColumn(csTab, "Variable Type", rowCount, columnFound)
    formatValues = ReadTableColumn(csTab, "Variable Format", rowCount, columnFound)
    uniqueValues = ReadTableColumn(csTab, "Unique", rowCount, columnFound)

    If Not hasMainLabel Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Main Label")
        check.Add "dict-mis-column-" & checkingCounter, infoMessage, checkingError
    End If

    If Not hasSheetName Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Sheet Name")
        check.Add "dict-mis-column-" & checkingCounter, infoMessage, checkingError
    End If

    If Not hasControl Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Control")
        check.Add "dict-mis-column-" & checkingCounter, infoMessage, checkingError
    End If

    Set nameCounts = CountNames(varValues, rowCount)
    choicesIndex = BuildIndex(this.choices.AllChoices())

    For rowIndex = rowCount To 1 Step -1
        worksheetRow = firstRow + rowIndex - 1
        checkingCounter = 0

        'The trimmed name answers the length test and the duplicate count. The
        'row itself carries every other value, so a name holding a double space
        'no longer costs the row seven skipped checks.
        varValue = workFun.Trim(varValues(rowIndex))
        controlValue = controlValues(rowIndex)
        controlDetailsValue = controlDetails(rowIndex)
        minValue = minValues(rowIndex)
        maxValue = maxValues(rowIndex)
        typeValue = typeValues(rowIndex)

        If NameCount(nameCounts, varValue) > 1 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-unique"
            infoMessage = FormatMessage(DICT_VAR_UNIQUE, worksheetRow, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
        End If

        If Len(varValue) < 4 Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-var-length"
            infoMessage = FormatMessage(DICT_VAR_LENGTH, worksheetRow, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
        End If

        If hasMainLabel Then
            If mainLabels(rowIndex) = vbNullString Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-main-lab"
                infoMessage = FormatMessage(DICT_MAIN_LAB, worksheetRow, varValue)
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
            End If
        End If

        If hasSheetName Then
            If sheetNames(rowIndex) = vbNullString Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-empty-sheet"
                infoMessage = FormatMessage(DICT_EMPTY_SHEET, worksheetRow, varValue)
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
            End If
        End If

        If (Not KnownControl(controlValue)) And _
           (controlValue <> vbNullString) And _
           (Not (InStr(1, controlValue, "choice_multiple") = 1)) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unknown-control"
            infoMessage = FormatMessage(DICT_UNKNOWN_CONTROL, worksheetRow, controlValue, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
        End If

        If (controlValue = "choice_formula") And (InStr(1, controlDetailsValue, "CHOICE_FORMULA") = 1) Then
            Set setupForm = ChoiceFormula.Create(controlDetailsValue)
            choiName = setupForm.ChoiceName()
            If Not IndexHolds(choicesIndex, choiName) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choiform-empty"
                infoMessage = FormatMessage(DICT_CHOIFORM_EMPTY, worksheetRow, choiName)
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
            Else
                keyName = "dict-cat-notfound"
                Set choiCategories = this.choices.Categories(choiName)
                Set formCategories = setupForm.Categories()

                For tabCounter = formCategories.LowerBound To formCategories.UpperBound
                    catValue = formCategories.Item(tabCounter)
                    If Not ContainsIgnoringCase(choiCategories, catValue) Then
                        checkingCounter = checkingCounter + 1
                        infoMessage = FormatMessage(DICT_CAT_NOTFOUND, worksheetRow, catValue, choiName)
                        check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingInfo
                    End If
                Next tabCounter
            End If
        End If

        If (controlValue = "formula" Or controlValue = "case_when" Or controlValue = "choice_formula") Then
            keyName = "dict-incor-form"
            infoMessage = FormulaMessage(controlDetailsValue, DICT_INCOR_FORM, worksheetRow, varValue)
            If infoMessage <> vbNullString Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (controlValue = "choice_manual") Or _
           ((controlValue = "choice_custom") And (controlDetailsValue <> vbNullString)) Or _
           (InStr(1, controlValue, "choice_multiple") = 1) Then
            If Not IndexHolds(choicesIndex, controlDetailsValue) Then
                checkingCounter = checkingCounter + 1
                keyName = "dict-choi-empty"
                infoMessage = FormatMessage(DICT_CHOI_EMPTY, worksheetRow, controlDetailsValue)
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (minValue <> vbNullString) Then
            keyName = "dict-incor-min"
            infoMessage = FormulaMessage(minValue, DICT_INCOR_MIN, worksheetRow, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (maxValue <> vbNullString) Then
            keyName = "dict-incor-max"
            infoMessage = FormulaMessage(maxValue, DICT_INCOR_MAX, worksheetRow, varValue)
            If (infoMessage <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If ((minValue <> vbNullString) Or (maxValue <> vbNullString)) And (typeValue = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-valid-control"
            infoMessage = FormatMessage(DICT_VALID_CONTROL, worksheetRow, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingError
        End If

        If (formatValues(rowIndex) <> vbNullString) And (typeValue = vbNullString) Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-format-control"
            infoMessage = FormatMessage(DICT_FORMAT_CONTROL, worksheetRow, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingWarning
        End If

        If (uniqueValues(rowIndex) = "yes") Then
            checkingCounter = checkingCounter + 1
            keyName = "dict-unique-val"
            infoMessage = FormatMessage(DICT_UNIQUE_VAL, worksheetRow, varValue)
            check.Add keyName & "-row-" & worksheetRow & "-" & checkingCounter, infoMessage, checkingInfo
        End If
    Next rowIndex

    expTab.Sort "Export Number"
    Set statusByNumber = ExportStatuses(expTab)

    For expCounter = 1 To this.dictionary.TotalNumberOfExports()
        Set expRng = csTab.DataRange("Export " & expCounter, strictSearch:=True)

        If Not (expRng Is Nothing) Then
            'Addressed by the number the Export Number cell carries, whatever
            'the text around it. LLExport fills that cell with "Export 1", so
            'a lookup by the bare "1" answered nothing and this check reported
            'nothing at all. An export number with no row in the Exports table
            'is still quiet.
            If ExportStatusOf(statusByNumber, expCounter, exportStatus) Then
                If (workFun.CountBlank(expRng) <> expRng.Rows.Count) And (exportStatus <> "active") Then
                    checkingCounter = checkingCounter + 1
                    keyName = "dict-export-na"
                    infoMessage = FormatMessage(DICT_EXPORT_NA, CStr(expCounter))
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
                End If
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

'@label:check-choice
'@sub-title Inspect choices sheet for unused or inconsistent entries
'@details
'Unprotects the choices sheet, builds a delimited index of the choice names
'actually referenced in dictionary controls, then flags unused choices,
'missing list names, missing ordering values, and missing labels. The three
'choices columns are read by name, so the check no longer depends on the
'distance between them. Collects all findings into a single Checking object
'and re-protects on exit.
Private Sub CheckChoice()
    Dim check As Checking
    Dim shChoices As Worksheet
    Dim shDictionary As Worksheet
    Dim choiTab As CustomTable
    Dim dictTab As CustomTable
    Dim choiList As BetterArray
    Dim usedChoicesList As BetterArray
    Dim usedIndex As String
    Dim counter As Long
    Dim checkingCounter As Long
    Dim choiName As String
    Dim infoMessage As String
    Dim keyName As String
    Dim setupForm As ChoiceFormula
    Dim choiNameRng As Range
    Dim controlRng As Range
    Dim controlDetailsRng As Range
    Dim actualControl As String
    Dim detailsValue As String
    Dim dictRowCount As Long
    Dim dictControls As Variant
    Dim dictDetails As Variant
    Dim choiRowCount As Long
    Dim choiFirstRow As Long
    Dim rowIndex As Long
    Dim worksheetRow As Long
    Dim choiNames As Variant
    Dim orderValues As Variant
    Dim labelValues As Variant
    Dim sortValue As String
    Dim choiLabValue As String
    Dim hasOrdering As Boolean
    Dim hasLabel As Boolean

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
    Set controlRng = dictTab.DataRange("Control", strictSearch:=True)
    Set controlDetailsRng = dictTab.DataRange("Control Details", strictSearch:=True)

    checkingCounter = 0

    If controlRng Is Nothing Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Control")
        check.Add "choi-mis-column-" & checkingCounter, infoMessage, checkingError
    ElseIf controlDetailsRng Is Nothing Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(DICT_MIS_COLUMN, "Control Details")
        check.Add "choi-mis-column-" & checkingCounter, infoMessage, checkingError
    Else
        dictRowCount = controlDetailsRng.Rows.Count
        dictControls = ColumnText(controlRng, dictRowCount)
        dictDetails = ColumnText(controlDetailsRng, dictRowCount)

        For counter = 1 To dictRowCount
            actualControl = dictControls(counter)
            detailsValue = dictDetails(counter)

            If (actualControl = "choice_manual") Or _
               (actualControl = "choice_custom") Or _
               (InStr(1, actualControl, "choice_multiple") = 1) Then

                If (detailsValue <> vbNullString) Then
                    usedChoicesList.Push detailsValue
                End If
            ElseIf (actualControl = "choice_formula") Then
                If (detailsValue <> vbNullString) Then
                    Set setupForm = ChoiceFormula.Create(detailsValue)
                    choiName = setupForm.ChoiceName()
                    If LenB(choiName) > 0 Then
                        usedChoicesList.Push choiName
                    End If
                End If
            End If
        Next counter
    End If

    usedIndex = BuildIndex(usedChoicesList)

    For counter = choiList.LowerBound To choiList.UpperBound
        choiName = CStr(choiList.Item(counter))
        If Not IndexHolds(usedIndex, choiName) Then
            checkingCounter = checkingCounter + 1
            keyName = "choi-unfound-choi"
            infoMessage = FormatMessage(CHOI_UNFOUND_CHOI, choiName)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
        End If
    Next counter

    Set choiNameRng = choiTab.DataRange("List Name", strictSearch:=True)

    If choiNameRng Is Nothing Then
        checkingCounter = checkingCounter + 1
        infoMessage = FormatMessage(CHOI_MIS_COLUMN, "List Name")
        check.Add "choi-mis-column-" & checkingCounter, infoMessage, checkingError
    Else
        choiFirstRow = choiNameRng.Row
        choiRowCount = choiNameRng.Rows.Count

        choiNames = ColumnText(choiNameRng, choiRowCount)
        orderValues = ReadTableColumn(choiTab, "Ordering list", choiRowCount, hasOrdering)
        labelValues = ReadTableColumn(choiTab, "Label", choiRowCount, hasLabel)

        If Not hasOrdering Then
            checkingCounter = checkingCounter + 1
            infoMessage = FormatMessage(CHOI_MIS_COLUMN, "Ordering list")
            check.Add "choi-mis-column-" & checkingCounter, infoMessage, checkingError
        End If

        If Not hasLabel Then
            checkingCounter = checkingCounter + 1
            infoMessage = FormatMessage(CHOI_MIS_COLUMN, "Label")
            check.Add "choi-mis-column-" & checkingCounter, infoMessage, checkingError
        End If

        For rowIndex = choiRowCount To 1 Step -1
            worksheetRow = choiFirstRow + rowIndex - 1
            choiName = choiNames(rowIndex)
            sortValue = orderValues(rowIndex)
            choiLabValue = labelValues(rowIndex)

            If (choiLabValue <> vbNullString) And (choiName = vbNullString) Then
                checkingCounter = checkingCounter + 1
                keyName = "choi-emptychoi-lab"
                infoMessage = FormatMessage(CHOI_EMPTYCHOI_LAB, worksheetRow, choiLabValue)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
            End If

            If (sortValue <> vbNullString) And (choiName = vbNullString) Then
                checkingCounter = checkingCounter + 1
                keyName = "choi-emptychoi-order"
                infoMessage = FormatMessage(CHOI_EMPTYCHOI_ORDER, worksheetRow, sortValue)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
            End If

            If (sortValue = vbNullString) And (choiName <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                keyName = "choi-empty-order"
                infoMessage = FormatMessage(CHOI_EMPTY_ORDER, worksheetRow, choiName)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
            End If

            If (choiLabValue = vbNullString) And (choiName <> vbNullString) Then
                checkingCounter = checkingCounter + 1
                keyName = "choi-mis-lab"
                infoMessage = FormatMessage(CHOI_MIS_LAB, worksheetRow, choiName)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        Next rowIndex
    End If

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

'@label:check-exports
'@sub-title Audit exports sheet for missing metadata and dictionary alignment
'@details
'Iterates over each export row, checking active exports for missing label,
'password, metadata, format, filename, and header values. Also validates
'that the corresponding dictionary column exists and is not empty, checks
'password/identifier consistency, and validates filename variable references.
'Every value of a row is read from that row, and the Export Number the row
'carries is what names the export in the message and picks the dictionary
'column. The old shape read the status by row position and the password by
'Export Number, so the two only agreed while the numbers ran 1..N in row
'order.
Private Sub CheckExports()

    Dim expTab As CustomTable
    Dim counter As Long
    Dim shExports As Worksheet
    Dim check As Checking
    Dim keyName As String
    Dim checkingCounter As Long
    Dim expStatus As String
    Dim infoMessage As String
    Dim headerNames As Variant
    Dim checkKeys As Variant
    Dim messageTexts As Variant
    Dim fieldRanges(0 To 5) As Range
    Dim headerCounter As Long
    Dim exportRng As Range
    Dim statusRng As Range
    Dim numberRng As Range
    Dim identifiersRng As Range
    Dim adminLevelsRng As Range
    Dim adminLevelsValue As String
    Dim workFun As WorksheetFunction
    Dim fileNameChunks As Variant
    Dim actualFileName As String
    Dim fileCounter As Long
    Dim fileNameChunk As String
    Dim vars As LLVariables
    Dim numberOfExports As Long
    Dim pwd As String
    Dim exportKeepsId As String
    Dim exportNumber As String
    Dim exportIdNumber As Long
    Dim worksheetRow As Long

    Set shExports = ResolveWorksheet(SHEET_EXPORTS)
    Set expTab = CustomTable.Create(shExports.ListObjects(1), idCol:="Export Number")
    Set check = Checking.Create("Export incoherences", "Type--Where?--Details")
    Set workFun = Application.WorksheetFunction
    Set vars = LLVariables.Create(this.dictionary)

    headerNames = Array("Label button", "Password", "Export metadata sheets", _
                        "File format", "File name", "Header format")
    checkKeys = Array("exp-mis-lab", "exp-mis-pass", "exp-mis-meta", _
                      "exp-mis-form", "exp-mis-name", "exp-mis-head")
    messageTexts = Array(EXP_MIS_LAB, EXP_MIS_PASS, EXP_MIS_META, _
                         EXP_MIS_FORM, EXP_MIS_NAME, EXP_MIS_HEAD)

    checkingCounter = 0

    Set statusRng = expTab.DataRange("Status", strictSearch:=True)

    If statusRng Is Nothing Then
        checkingCounter = checkingCounter + 1
        check.Add "exp-mis-status-" & checkingCounter, EXP_MIS_STATUS, checkingError
        this.checkTables.Push check
        Exit Sub
    End If

    Set numberRng = expTab.DataRange("Export Number", strictSearch:=True)
    Set identifiersRng = expTab.DataRange("Include personal identifiers", strictSearch:=True)

    'The column is optional, so a setup written before it existed is not reported
    'as missing a column. An export whose row has no cell to read exports every
    'admin level, which is what LLExporter.KeptAdminLevels answers for it.
    Set adminLevelsRng = expTab.DataRange("Admin levels", strictSearch:=True)

    'The six required columns are resolved once and reported once. Reading them
    'cell by cell answered Nothing for a missing column, and IsEmpty(Nothing) is
    'False, so a missing column used to read as "this field is filled".
    For headerCounter = 0 To 5
        Set fieldRanges(headerCounter) = expTab.DataRange(CStr(headerNames(headerCounter)), strictSearch:=True)
        If fieldRanges(headerCounter) Is Nothing Then
            checkingCounter = checkingCounter + 1
            infoMessage = FormatMessage(EXP_MIS_COLUMN, CStr(headerNames(headerCounter)))
            check.Add "exp-mis-column-" & checkingCounter, infoMessage, checkingError
        End If
    Next headerCounter

    numberOfExports = statusRng.Rows.Count

    For counter = 1 To numberOfExports

        worksheetRow = statusRng.Row + counter - 1
        expStatus = CellText(statusRng.Cells(counter, 1).Value)

        'The cell reads "Export 1", so the number is parsed out of it before
        'anything is built from it. Reading the cell whole and prepending
        '"Export " again asked the dictionary for "Export Export 1" and
        'reported every active export as missing its column.
        exportIdNumber = 0
        If Not numberRng Is Nothing Then _
            exportIdNumber = ExportIdentifierNumber(numberRng.Cells(counter, 1).Value)
        If exportIdNumber <= 0 Then exportIdNumber = counter
        exportNumber = CStr(exportIdNumber)

        pwd = vbNullString
        If Not fieldRanges(1) Is Nothing Then pwd = CellText(fieldRanges(1).Cells(counter, 1).Value)

        exportKeepsId = vbNullString
        If Not identifiersRng Is Nothing Then exportKeepsId = CellText(identifiersRng.Cells(counter, 1).Value)

        adminLevelsValue = vbNullString
        If Not adminLevelsRng Is Nothing Then _
            adminLevelsValue = CellText(adminLevelsRng.Cells(counter, 1).Value)

        For headerCounter = 0 To 5
            If Not fieldRanges(headerCounter) Is Nothing Then
                If (CellText(fieldRanges(headerCounter).Cells(counter, 1).Value) = vbNullString) And _
                   (expStatus = "active") Then
                    checkingCounter = checkingCounter + 1
                    keyName = CStr(checkKeys(headerCounter))
                    infoMessage = FormatMessage(CStr(messageTexts(headerCounter)), worksheetRow, exportNumber)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
                End If
            End If
        Next headerCounter

        On Error Resume Next
            Set exportRng = Nothing
            Set exportRng = this.dictionary.DataRange("Export " & exportNumber, includeHeaders:=False)
        On Error GoTo 0

        If (expStatus = "active") And (exportRng Is Nothing) Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-unfound-dictcolumn"
            infoMessage = FormatMessage(EXP_UNFOUND_DICTCOLUMN, worksheetRow, exportNumber)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        ElseIf (expStatus = "active") And (Not exportRng Is Nothing) Then
            If workFun.CountBlank(exportRng) = exportRng.Rows.Count Then
                checkingCounter = checkingCounter + 1
                keyName = "exp-act-empty"
                infoMessage = FormatMessage(EXP_ACT_EMPTY, worksheetRow, exportNumber)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
            End If
        End If

        If (pwd <> "yes") And (exportKeepsId = "yes") Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-id-passwd"
            infoMessage = FormatMessage(EXP_ID_PASSWD, worksheetRow, exportNumber)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
        End If

        'A ROW OF ANY STATUS IS READ HERE
        '-----------------------------------------------------------------------
        'The row above reports what an ACTIVE export is missing, because an
        'inactive one writes no file. This reports a cell that is written wrong,
        'and a row turned active later carries the same wrong cell with it.
        If Not IsAdminLevelsCell(adminLevelsValue) Then
            checkingCounter = checkingCounter + 1
            keyName = "exp-adm-levels"
            infoMessage = FormatMessage(EXP_ADM_LEVELS, worksheetRow, exportNumber, _
                                        adminLevelsValue)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
        End If

        actualFileName = vbNullString
        If Not fieldRanges(4) Is Nothing Then actualFileName = CellText(fieldRanges(4).Cells(counter, 1).Value)

        'An empty file name is already reported by the loop above. Splitting it
        'gives one empty chunk, which used to be read as a variable named "".
        If LenB(actualFileName) > 0 Then
            fileNameChunks = Split(actualFileName, "+")

            For fileCounter = LBound(fileNameChunks) To UBound(fileNameChunks)
                'The three rules below are LLExport.ResolveTemplateChunk's own:
                'the chunk is trimmed, a literal starts and ends with a quote,
                'and the sheet type is compared without regard to case.
                fileNameChunk = Trim$(CStr(fileNameChunks(fileCounter)))

                If LenB(fileNameChunk) > 0 Then
                    If Not IsLiteralChunk(fileNameChunk) Then
                        If Not vars.Contains(fileNameChunk) Then
                            checkingCounter = checkingCounter + 1
                            keyName = "exp-filename-varunfound"
                            infoMessage = FormatMessage(EXP_FILENAME_VARUNFOUND, worksheetRow, exportNumber, fileNameChunk)
                            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                        ElseIf StrComp(vars.Value(colName:="Sheet Type", varName:=fileNameChunk), _
                                       "vlist1D", vbTextCompare) <> 0 Then
                            checkingCounter = checkingCounter + 1
                            keyName = "exp-filename-varnotvlist"
                            infoMessage = FormatMessage(EXP_FILENAME_VARNOTVLIST, worksheetRow, exportNumber, fileNameChunk)
                            check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                        End If
                    End If
                End If
            Next fileCounter
        End If
    Next counter

    this.checkTables.Push check
End Sub

'@label:check-translations
'@sub-title Review translations table for missing and duplicate labels
'@details
'Unprotects the translations sheet, updates the translations manager from
'the registry, then queries for missing and duplicate labels. Collects
'findings into a single Checking object and re-protects on exit.
Private Sub CheckTranslations()

    Dim check As Checking
    Dim counter As Long
    Dim summary As BetterArray
    Dim shRegistry As Worksheet

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo TranslationsFailed

    Set shRegistry = ResolveWorksheet(SHEET_REGISTRY)
    Set check = Checking.Create("Translations", "Incoherence--Where?--Details")

    'unlock the translation worksheet before updating
    this.passwordService.UnProtect SHEET_TRANSLATIONS
    this.transManager.SetDisplayPrompts False
    this.transManager.UpdateFromRegistry shRegistry

    'Missing labels count
    Set summary = this.transManager.InternalMissingSummary()

    For counter = summary.LowerBound To summary.UpperBound
        check.Add "trads-mis-labs-" & counter, summary.Item(counter), checkingInfo
    Next

    'Duplicate labels count
    Set summary = this.transManager.InternalDuplicatesSummary()

    For counter = summary.LowerBound To summary.UpperBound
        check.Add "trads-dup-labs-" & counter, summary.Item(counter), checkingWarning
    Next
    this.checkTables.Push check

    this.passwordService.Protect SHEET_TRANSLATIONS
    this.transManager.SetDisplayPrompts True
    Exit Sub

TranslationsFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_TRANSLATIONS
        this.transManager.SetDisplayPrompts True
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@label:check-analysis
'@sub-title Run analysis table checks across all sections
'@details
'Resolves the analysis worksheet, removes empty rows through
'TidyAnalysisSheet, reports a failure there as a Checking entry, then
'delegates to CheckTable for each of the seven analysis sections.
'Caches the worksheet and WorksheetFunction references and passes them to
'CheckTable to avoid repeated lookups.
Private Sub CheckAnalysis()
    Dim analysisSheet As Worksheet
    Dim tidyCheck As Checking
    Dim failureReason As String
    Dim selectionCell As Range
    Dim previousSelection As String
    Dim selectionConfigured As Boolean
    Dim workFun As WorksheetFunction

    Set analysisSheet = ResolveWorksheet(SHEET_ANALYSIS)

    'RNG_SelectTable is a defined name a setup workbook may not carry, so a
    'failure to resolve it means "there is nothing to drive here".
    On Error Resume Next
        Set selectionCell = analysisSheet.Range("RNG_SelectTable")
        If Not selectionCell Is Nothing Then
            selectionConfigured = True
            previousSelection = CStr(selectionCell.Value)
            selectionCell.Value = "Add or remove rows of all tables"
        End If
    On Error GoTo 0

    failureReason = TidyAnalysisSheet(analysisSheet)

    If LenB(failureReason) > 0 Then
        Set tidyCheck = Checking.Create("Analysis incoherences", _
                                        "Analysis sheet--Where?--Details")
        tidyCheck.Add "ana-tidy-failed-1", _
                      FormatMessage(ANA_TIDY_FAILED, failureReason), checkingError
        this.checkTables.Push tidyCheck
    End If

    If selectionConfigured Then
        On Error Resume Next
            selectionCell.Value = previousSelection
        On Error GoTo 0
    End If

    Set workFun = Application.WorksheetFunction

    CheckTable "Global summary", analysisSheet, workFun
    CheckTable "Univariate analysis", analysisSheet, workFun
    CheckTable "Bivariate analysis", analysisSheet, workFun
    CheckTable "Time series analysis", analysisSheet, workFun
    CheckTable "Time series graphs", analysisSheet, workFun
    CheckTable "Spatial analysis", analysisSheet, workFun
    CheckTable "Spatio-temporal analysis", analysisSheet, workFun
End Sub

'@label:tidy-analysis-sheet
'@sub-title Remove the empty analysis rows before the checks read the sheet
'@details
'The sheet is unprotected, trimmed and protected again. When any of the three
'fails, the sheet is protected in the handler and the reason is handed back so
'the caller can report it. Passwords.UnProtect stays quiet when the stored key
'does not open the sheet, and the row removal that follows then raises on a
'sheet that is still locked, so this failure is a real one for a user.
'@param analysisSheet Worksheet. The analysis worksheet.
'@return String. The failure description, or an empty string on success.
Private Function TidyAnalysisSheet(ByVal analysisSheet As Worksheet) As String
    Dim analysisManager As Analysis

    On Error GoTo TidyFailed

    this.passwordService.UnProtect SHEET_ANALYSIS
    Set analysisManager = Analysis.Create(analysisSheet)
    analysisManager.RemoveRows
    this.passwordService.Protect SHEET_ANALYSIS

    On Error GoTo 0
    Exit Function

TidyFailed:
    TidyAnalysisSheet = Err.Description
    On Error Resume Next
        this.passwordService.Protect SHEET_ANALYSIS
    On Error GoTo 0
End Function

'@sub-title Evaluate a specific analysis section for structural errors.
'@param partName String. The section display name (e.g. "Global summary").
'@param analysisSheet Worksheet. The cached analysis worksheet.
'@param workFun WorksheetFunction. The cached worksheet function object.
Private Sub CheckTable(ByVal partName As String, _
                       ByVal analysisSheet As Worksheet, _
                       ByVal workFun As WorksheetFunction)

    Const TAB_GLOBAL_SUMMARY As String = "Tab_Global_Summary"
    Const TAB_UNIVARIATE As String = "Tab_Univariate_Analysis"
    Const TAB_BIVARIATE As String = "Tab_Bivariate_Analysis"
    Const TAB_TIME_SERIES As String = "Tab_TimeSeries_Analysis"
    Const TAB_SPATIAL As String = "Tab_Spatial_Analysis"
    Const TAB_SPATIOTEMP As String = "Tab_SpatioTemporal_Analysis"
    Const TAB_TIME_SERIES_GRAPHS As String = "Tab_Graph_TimeSeries"

    Dim listTable As ListObject
    Dim listName As String
    Dim headerRange As Range
    Dim specsRange As Range
    Dim specs As TableSpecs
    Dim specsError As String
    Dim check As Checking
    Dim rowIndex As Long
    Dim checkingCounter As Long
    Dim keyName As String
    Dim infoMessage As String
    Dim analysisFormula As String
    Dim tableScope As Byte
    Dim blankCount As Long
    Dim columnCount As Long

    listName = Switch( _
        partName = "Global summary", TAB_GLOBAL_SUMMARY, _
        partName = "Univariate analysis", TAB_UNIVARIATE, _
        partName = "Bivariate analysis", TAB_BIVARIATE, _
        partName = "Time series analysis", TAB_TIME_SERIES, _
        partName = "Spatial analysis", TAB_SPATIAL, _
        partName = "Spatio-temporal analysis", TAB_SPATIOTEMP, _
        partName = "Time series graphs", TAB_TIME_SERIES_GRAPHS _
    )
    checkingCounter = 0

    On Error Resume Next
        Set listTable = analysisSheet.ListObjects(listName)
    On Error GoTo 0
    If listTable Is Nothing Then Exit Sub

    Set check = Checking.Create("Analysis incoherences", _
                                partName & "--Where?--Details")

    Set headerRange = listTable.HeaderRowRange

    For rowIndex = 1 To listTable.ListRows.Count
        Set specsRange = listTable.ListRows(rowIndex).Range
        blankCount = workFun.CountBlank(specsRange)
        columnCount = specsRange.Columns.Count

        If ((blankCount = columnCount) And (listName <> TAB_TIME_SERIES)) Or _
           (blankCount >= columnCount - 1) Then

            keyName = "ana-empty-tab"
            checkingCounter = checkingCounter + 1
            infoMessage = FormatMessage(ANA_EMPTY_TAB, specsRange.Row)
            check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
        Else

            Set specs = Nothing
            specsError = vbNullString
            Err.Clear
            On Error Resume Next
            Set specs = TableSpecs.Create(headerRange, specsRange, this.dictionary)
            specsError = Err.Description
            On Error GoTo 0
            Err.Clear

            'A row TableSpecs refuses to read is a row the analysis build will
            'refuse too, so it is reported rather than passed over in silence.
            If specs Is Nothing Then
                checkingCounter = checkingCounter + 1
                keyName = "ana-specs-failed"
                infoMessage = FormatMessage(ANA_SPECS_FAILED, specsRange.Row, specsError)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
                GoTo ContinueRow
            End If

            tableScope = specs.TableScope

            If (Not specs.ValidTable()) Then
                checkingCounter = checkingCounter + 1
                keyName = "ana-inv-tab"
                infoMessage = FormatMessage(ANA_INV_TAB, specsRange.Row, "check row and column variables")
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingError

                If (tableScope = ScopeTimeSeries) And (specs.IsNewSection()) _
                   And (blankCount < columnCount - 1) Then
                    keyName = "ana-ts-newsec"
                    infoMessage = FormatMessage(ANA_TS_NEWSEC, specsRange.Row)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingError
                End If
            End If

            If (specs.Value("title") = vbNullString) And _
               (tableScope <> ScopeGlobalSummary) And _
               (tableScope <> ScopeTimeSeriesGraph) Then

                checkingCounter = checkingCounter + 1
                keyName = "ana-empty-title"
                infoMessage = FormatMessage(ANA_EMPTY_TITLE, specsRange.Row)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
            End If

            If (tableScope = ScopeUnivariate) And _
               specs.HasGraph() And _
               specs.FlipCoordinates() And _
               (specs.Value("percentage") = "yes") Then

                checkingCounter = checkingCounter + 1
                keyName = "ana-uaflip-perc"
                infoMessage = FormatMessage(ANA_UAFLIP_PERC, specsRange.Row)
                check.Add keyName & "-" & checkingCounter, infoMessage, checkingNote
            End If

            If (tableScope <> ScopeTimeSeries) And (tableScope <> ScopeTimeSeriesGraph) Then
                If (Not specs.HasGraph()) And specs.FlipCoordinates() Then
                    checkingCounter = checkingCounter + 1
                    keyName = "ana-adgr-flip"
                    infoMessage = FormatMessage(ANA_ADGR_FLIP, specsRange.Row)
                    check.Add keyName & "-" & checkingCounter, infoMessage, checkingInfo
                End If
            End If

            If (tableScope <> ScopeTimeSeriesGraph) Then
                analysisFormula = specs.Value("function")
                If (analysisFormula <> vbNullString) Then
                    keyName = "ana-incor-form"
                    infoMessage = FormulaMessage(analysisFormula, ANA_INCOR_FORM, specsRange.Row, _
                                                 formulaType:=CONTEXT_ANALYSIS)
                    If (infoMessage <> vbNullString) Then
                        checkingCounter = checkingCounter + 1
                        check.Add keyName & "-" & checkingCounter, infoMessage, checkingWarning
                    End If
                End If
            End If
        End If
ContinueRow:
    Next rowIndex

    this.checkTables.Push check
End Sub

'@label:print-report
'@sub-title Publish all collected checkings to the reporting worksheet
'@details
'Resolves the __checkRep worksheet, clears its content, creates a
'CheckingOutput helper, and prints all collected Checking objects.
'Installs the worksheet change handler for filtering support and
'activates the report sheet.
'The clear needs no unprotect: __checkRep is deliberately outside the
'protection matrix and CheckingOutput unprotects it before it writes anyway.
'Only the change handler is allowed to fail quietly, because it writes into
'the VBA project and many machines block that. Activating the report is what
'the user pressed the button for, so a failure there is reported.
Private Sub PrintReport()
    Dim checkOutput As CheckingOutput
    Dim reportSheet As Worksheet

    Set reportSheet = ResolveWorksheet(SHEET_CHECK_OUTPUT)
    reportSheet.Cells.Clear
    Set checkOutput = CheckingOutput.Create(reportSheet)

    reportSheet.Cells.EntireRow.Hidden = False

    checkOutput.PrintOutput this.checkTables

    On Error Resume Next
        checkOutput.EnsureWorksheetChangeHandler
    On Error GoTo 0

    reportSheet.Activate
End Sub

'@section Message helpers
'===============================================================================

'@label:format-message
'@sub-title Format an error message by replacing numbered placeholders
'@param messageTag String. The message template with %1..%4 placeholders.
'@param valueOne Optional String. Replacement for %1.
'@param valueTwo Optional String. Replacement for %2.
'@param valueThree Optional String. Replacement for %3.
'@param valueFour Optional String. Replacement for %4.
'@return String. The formatted message.
Private Function FormatMessage(ByVal messageTag As String, _
                               Optional ByVal valueOne As String = vbNullString, _
                               Optional ByVal valueTwo As String = vbNullString, _
                               Optional ByVal valueThree As String = vbNullString, _
                               Optional ByVal valueFour As String = vbNullString) As String

    Dim template As String

    template = messageTag

    'One test answers for all four passes on a template that carries no
    'placeholder at all.
    If InStr(1, template, "%", vbBinaryCompare) = 0 Then
        FormatMessage = template
        Exit Function
    End If

    template = ReplacePlaceholder(template, "%1", valueOne)
    template = ReplacePlaceholder(template, "%2", valueTwo)
    template = ReplacePlaceholder(template, "%3", valueThree)
    template = ReplacePlaceholder(template, "%4", valueFour)

    FormatMessage = template
End Function

'@label:formula-checker
'@sub-title The one formula parser a run reuses
'@details
'Built with New rather than through Formulas.Create, because the factory seals
'the instance and a sealed instance refuses SetupFormula. Reuse is what keeps
'the LLVariables cache warm: a fresh parser per formula read the whole
'Variable Name column again, up to three times per dictionary row.
'@return Formulas. The parser bound to this run's dictionary and formula store.
Private Function FormulaChecker() As Formulas
    If this.formulaChecker Is Nothing Then
        Set this.formulaChecker = New Formulas
        this.formulaChecker.Initialise this.dictionary, this.formulaStore, vbNullString
    End If

    Set FormulaChecker = this.formulaChecker
End Function

'@label:formula-message
'@sub-title Validate a formula string and return a formatted error when invalid
'@details
'The reason is read in the context the formula was validated in. Asking for
'the default context instead made a formula that is invalid as a linelist
'formula and valid as an analysis one report "The formula seems correct" at
'the end of the line saying it will fail.
'@param formValue String. The formula expression to validate.
'@param messageTag String. The message template to use on failure.
'@param valueOne Optional String. Replacement for %1.
'@param valueTwo Optional String. Replacement for %2.
'@param formulaType Optional String. Formula type ("linelist" or "analysis"). Defaults to "linelist".
'@return String. The formatted error message, or empty when the formula is valid.
Private Function FormulaMessage(ByVal formValue As String, _
                                ByVal messageTag As String, _
                                Optional ByVal valueOne As String = vbNullString, _
                                Optional ByVal valueTwo As String = vbNullString, _
                                Optional ByVal formulaType As String = CONTEXT_LINELIST) As String
    Dim setupFormulas As Formulas

    If LenB(Trim$(formValue)) = 0 Then Exit Function

    On Error GoTo FormulaFail
        Set setupFormulas = FormulaChecker()
        setupFormulas.SetupFormula = formValue
        If Not setupFormulas.Valid(formulaType:=formulaType) Then
            FormulaMessage = FormatMessage(messageTag, valueOne, valueTwo, _
                                           setupFormulas.Reason(formulaType:=formulaType))
        End If
    On Error GoTo 0
    Exit Function

FormulaFail:
    ' If validation fails unexpectedly, surface the failure as a formatted message.
    FormulaMessage = FormatMessage(messageTag, valueOne, valueTwo, Err.Description)
    Err.Clear
End Function

'@label:replace-placeholder
'@sub-title Replace a single placeholder token in a template string
'@param template String. The template containing the placeholder.
'@param placeholder String. The token to replace (e.g. "%1").
'@param replacement String. The value to substitute.
'@return String. The template with the placeholder replaced.
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

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber ProjectError. The error code to raise.
'@param errorMessage String. Human-readable description of the failure.
'@throws ProjectError Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASS_NAME, "Error: " & errorMessage
End Sub
