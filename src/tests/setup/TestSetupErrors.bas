Attribute VB_Name = "TestSetupErrors"
Attribute VB_Description = "Verifies SetupErrors orchestrator initialisation"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FORMULA_SHEET As String = "__formula"
Private Const PASSWORD_SHEET As String = "__pass"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const EXPORTS_SHEET As String = "Exports"
Private Const ANALYSIS_SHEET As String = "Analysis"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DROPDOWN_SHEET As String = "__variables"
Private Const CHECK_OUTPUT_SHEET As String = "__checkRep"
Private Const REGISTRY_SHEET As String = "__updated"

'Dictionary geometry of the fixture: headers on row 5, data rows 6 to 12.
Private Const DICT_HEADER_ROW As Long = 5
Private Const DICT_LAST_DATA_ROW As Long = 12

'A password the Passwords service does not hold, used to make the Analysis
'tidy-up fail the way a real locked sheet makes it fail.

'A file name chunk that resolves to a literal, so a test can silence every
'export it is not looking at.
Private Const SAFE_FILE_NAME As String = """report"""


'@Folder("CustomTests")
'@Folder("Tests.Setup")
'@ModuleDescription("Verifies SetupErrors orchestrator initialisation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest

'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupErrors"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Tests
'===============================================================================

'@TestMethod("SetupErrors")
Public Sub TestCreateReturnsInterface()
    CustomTestSetTitles Assert, "SetupErrors", "TestCreateReturnsInterface"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors

    Set hostBook = PrepareSetupWorkbook()
    Set checker = SetupErrors.Create(hostBook)

    Assert.IsNotNothing checker, "Factory should return a SetupErrors instance"

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    Set checker = Nothing
    CustomTestLogFailure Assert, "TestCreateReturnsInterface", Err.Number, Err.Description
    DeleteWorkbook hostBook
End Sub

'@TestMethod("SetupErrors")
Public Sub TestCheckingsInitialisedEmpty()
    CustomTestSetTitles Assert, "SetupErrors", "TestCheckingsInitialisedEmpty"
    On Error GoTo Fail

    Dim results As BetterArray
    Dim hostBook As Workbook
    Dim checker As SetupErrors

    Set hostBook = PrepareSetupWorkbook()
    Set checker = SetupErrors.Create(hostBook)

    Set results = checker.Checkings

    Assert.IsNotNothing results, "Checkings container should be initialised during setup"
    Assert.AreEqual 0&, results.Length, "Expected empty checkings before running the workflow"

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCheckingsInitialisedEmpty", Err.Number, Err.Description
    DeleteWorkbook hostBook
    Set checker = Nothing
End Sub

'@TestMethod("SetupErrors")
Public Sub TestCreateRaisesWithoutWorkbook()
    CustomTestSetTitles Assert, "SetupErrors", "TestCreateRaisesWithoutWorkbook"
    On Error GoTo Fail

    On Error GoTo ExpectMissingWorkbook
        SetupErrors.Create Nothing
        On Error GoTo Fail
        Assert.LogFailure "Create should raise when workbook reference is missing."
        Exit Sub

ExpectMissingWorkbook:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "Expected ObjectNotInitialized when workbook argument is Nothing."
    Err.Clear
    On Error GoTo Fail
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateRaisesWithoutWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("SetupErrors")
Public Sub TestRunRaisesWhenNotInitialised()
    CustomTestSetTitles Assert, "SetupErrors", "TestRunRaisesWhenNotInitialised"
    On Error GoTo Fail

    Dim subject As SetupErrors
    Set subject = New SetupErrors

    On Error GoTo ExpectNotInitialised
        subject.Run
        On Error GoTo Fail
        Assert.LogFailure "Run must raise when Initialise has not been called."
        Exit Sub

ExpectNotInitialised:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "Expected ObjectNotInitialized when calling Run without initialising."
    Err.Clear
    On Error GoTo Fail
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRunRaisesWhenNotInitialised", Err.Number, Err.Description
End Sub

'@TestMethod("SetupErrors")
'Covers C18. Against the old code Initialise was Public with no guard, so a
'holder could re-point a live checker at another workbook between Run and
'Checkings.
Public Sub TestInitialiseRaisesOnASealedChecker()
    CustomTestSetTitles Assert, "SetupErrors", "TestInitialiseRaisesOnASealedChecker"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors

    Set hostBook = PrepareSetupWorkbook()
    Set checker = SetupErrors.Create(hostBook)

    On Error GoTo ExpectSealed
        checker.Initialise hostBook
        On Error GoTo Fail
        Assert.LogFailure "Initialise should raise once the factory has sealed the instance."
        GoTo Cleanup

ExpectSealed:
    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), CLng(Err.Number), _
                    "Expected SomethingWentWrong when Initialise runs on a sealed checker."
    Err.Clear

Cleanup:
    On Error Resume Next
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    Set checker = Nothing
    CustomTestLogFailure Assert, "TestInitialiseRaisesOnASealedChecker", Err.Number, Err.Description
    DeleteWorkbook hostBook
End Sub

'@TestMethod("SetupErrors")
'Covers C2 as far as an assertion can reach it: a failure inside Run leaves the
'application settings as they were. The raise that happens INSIDE
'EnterBusyState is closed by arming the handler above it, and no test can force
'ApplyBusyState to fail from here.
Public Sub TestRunRestoresApplicationStateOnFailure()
    CustomTestSetTitles Assert, "SetupErrors", "TestRunRestoresApplicationStateOnFailure"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim subject As SetupErrors
    Dim previousCalcBeforeSave As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousAlerts As Boolean

    previousCalcBeforeSave = Application.CalculateBeforeSave
    previousScreenUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts
    Set hostBook = PrepareSetupWorkbook()
    Set subject = SetupErrors.Create(hostBook)

    Application.DisplayAlerts = False
    hostBook.Worksheets(CHECK_OUTPUT_SHEET).Delete
    Application.DisplayAlerts = previousAlerts

    On Error GoTo ExpectMissingWorksheet
        subject.Run
        On Error GoTo Fail
        Assert.LogFailure "Run should bubble when required worksheets are absent."
        GoTo CleanupFailure

ExpectMissingWorksheet:
    Assert.AreEqual CLng(ProjectError.ElementNotFound), CLng(Err.Number), _
                    "Run should raise ElementNotFound when the checking output sheet is missing."
    Err.Clear
    Assert.AreEqual previousCalcBeforeSave, Application.CalculateBeforeSave, _
                    "CalculateBeforeSave should be restored after failure."
    Assert.AreEqual previousScreenUpdating, Application.ScreenUpdating, _
                    "ScreenUpdating should be restored after failure."
    Assert.IsNotNothing subject.Checkings, _
                        "Failure should not make the checkings container unusable."

CleanupFailure:
    Application.DisplayAlerts = previousAlerts
    Set subject = Nothing
    DeleteWorkbook hostBook
    Application.CalculateBeforeSave = previousCalcBeforeSave
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

Fail:
    Application.DisplayAlerts = previousAlerts
    Set subject = Nothing
    Application.CalculateBeforeSave = previousCalcBeforeSave
    Application.ScreenUpdating = previousScreenUpdating
    CustomTestLogFailure Assert, "TestRunRestoresApplicationStateOnFailure", Err.Number, Err.Description
    DeleteWorkbook hostBook
End Sub

'@TestMethod("SetupErrors")
'Covers C3. Against the old code LeaveBusyState called Restore whatever
'happened, so a scope the caller had already made busy came back with screen
'updating on, calculation automatic and events on the moment Run returned.
Public Sub TestRunLeavesInjectedBusyScopeAlone()
    CustomTestSetTitles Assert, "SetupErrors", "TestRunLeavesInjectedBusyScopeAlone"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim scope As ApplicationState

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set checker = SetupErrors.Create(hostBook, stateScope:=scope)
    checker.Run

    Assert.IsTrue scope.IsBusy, _
                  "A busy scope handed in by the caller should still be busy after Run."

Cleanup:
    On Error Resume Next
    scope.Restore silent:=True
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    On Error Resume Next
    scope.Restore silent:=True
    Set checker = Nothing
    CustomTestLogFailure Assert, "TestRunLeavesInjectedBusyScopeAlone", Err.Number, Err.Description
    DeleteWorkbook hostBook
End Sub

'@TestMethod("SetupErrors")
Public Sub TestRunDetectsDictionaryAndChoicesIssues()
    CustomTestSetTitles Assert, "SetupErrors", "TestRunDetectsDictionaryAndChoicesIssues"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray
    Dim duplicateDetected As Boolean
    Dim missingLabelDetected As Boolean

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run

    Set results = checker.Checkings
    duplicateDetected = CheckingsContain(results, "Variable ""dup_variable"" is duplicate")
    missingLabelDetected = CheckingsContain(results, "missing Label for choice ""list_primary""")

    Assert.IsTrue duplicateDetected, _
                   "Dictionary checks should report duplicate variables."
    Assert.IsTrue missingLabelDetected, _
                   "Choices checks should report missing labels."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    Set checker = Nothing
    CustomTestLogFailure Assert, "TestRunDetectsDictionaryAndChoicesIssues", Err.Number, Err.Description
    DeleteWorkbook hostBook
End Sub

'@TestMethod("SetupErrors")
Public Sub TestDictionaryChecksReportAllMessages()
    CustomTestSetTitles Assert, "SetupErrors", "TestDictionaryChecksReportAllMessages"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray
    Dim expected As Variant
    Dim idx As Long

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    expected = Array( _
        "Variable ""dup_variable"" is duplicate", _
        "main label of variable ""dup_variable"" is empty", _
        "Sheet names should not be empty", _
        "choice manual ""missing_choice"" is not present", _
        "choice formula ""missing_choice"" is not present", _
        "category ""Unknown option"" does not exists for choice ""list_primary""", _
        "control ""unknown_control"" of variable ""unknown_ctrl"" is unknown", _
        "formula for variable ""calc_invalid"" will eventually fail", _
        "Max validation for variable ""calc_invalid""", _
        "Min validation for variable ""calc_invalid""", _
        "Duplicate values of variable dup_variable", _
        "Validation requires the type of the variable, please consider adding the type for variable calc_invalid", _
        "Format requires the type of the variable, please consider adding the type for variable calc_invalid", _
        "Variable ""abc"" length is less than 4", _
        """Export Number"" 1 status is not active")

    For idx = LBound(expected) To UBound(expected)
        AssertContainsMessage results, CStr(expected(idx)), "Dictionary"
    Next idx

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDictionaryChecksReportAllMessages", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
Public Sub TestChoicesChecksReportAllMessages()
    CustomTestSetTitles Assert, "SetupErrors", "TestChoicesChecksReportAllMessages"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray
    Dim expected As Variant
    Dim idx As Long

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    expected = Array( _
        "There is no attached ""List Name"" value for Label ""Orphan label""", _
        "There is no attached ""List Name"" value for Ordering list ""5""", _
        "There is no value of ""Ordering list"" for choice ""list_missing_order""", _
        "There is a missing Label for choice ""list_primary""", _
        "Choice name ""list_secondary"" is declared in choices sheet but never used")

    For idx = LBound(expected) To UBound(expected)
        AssertContainsMessage results, CStr(expected(idx)), "Choices"
    Next idx

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestChoicesChecksReportAllMessages", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
Public Sub TestExportsChecksReportAllMessages()
    CustomTestSetTitles Assert, "SetupErrors", "TestExportsChecksReportAllMessages"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray
    Dim expected As Variant
    Dim idx As Long

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    expected = Array( _
        "The Export Number 2 is active, but there is no label attached", _
        "The Export Number 2 is active, but there is no value for ""Password""", _
        "The Export Number 2 is active, but there is no value for ""Export Metadata""", _
        "The Export Number 2 is active, but there is no value for ""File format""", _
        "The Export Number 2 is active, but there is no value for ""File name""", _
        "The Export Number 2 is active, but there is no value for ""Header format""", _
        "has no password, but keeps identifiers", _
        "corresponding column is empty in the dictionary", _
        "file name contains a variable named ""missing_var"" which does not exists", _
        "file name contains a variable named ""calc_invalid"" which is not a vlist1D variable", _
        "corresponding column in the dictionary is missing")

    For idx = LBound(expected) To UBound(expected)
        AssertContainsMessage results, CStr(expected(idx)), "Exports"
    Next idx

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportsChecksReportAllMessages", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
Public Sub TestAnalysisChecksDetectInvalidTables()
    CustomTestSetTitles Assert, "SetupErrors", "TestAnalysisChecksDetectInvalidTables"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "The table is invalid"), _
                   "Analysis checks should report invalid tables."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAnalysisChecksDetectInvalidTables", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
Public Sub TestAnalysisChecksDetectEmptyRows()
    CustomTestSetTitles Assert, "SetupErrors", "TestAnalysisChecksDetectEmptyRows"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)

    AddEmptyAnalysisRow hostBook

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "This line is completely empty"), _
                   "Analysis checks should report empty table rows."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAnalysisChecksDetectEmptyRows", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
Public Sub TestAnalysisChecksProduceCheckingObject()
    CustomTestSetTitles Assert, "SetupErrors", "TestAnalysisChecksProduceCheckingObject"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue results.Length >= 5, _
                   "Run should produce at least 5 checking objects (dict, choices, exports, analysis, translations). Actual: " & results.Length

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAnalysisChecksProduceCheckingObject", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers the whole reporting path, which no assertion reached before: Run has to
'leave the findings on the __checkRep worksheet.
Public Sub TestRunPublishesTheReportSheet()
    CustomTestSetTitles Assert, "SetupErrors", "TestRunPublishesTheReportSheet"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run

    Assert.IsTrue SheetHoldsText(hostBook.Worksheets(CHECK_OUTPUT_SHEET), "incoherences"), _
                  "Run should print the checking titles on the report worksheet."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRunPublishesTheReportSheet", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C1. N() is valid as an analysis formula and invalid as a linelist one.
'The old code asked Reason for the default analysis context, so the line saying
'the formula will fail ended with "The formula seems correct". MEAN cannot play
'this part here: the fixture worksheet lists it among the Excel functions, so
'the tokeniser takes it in both contexts.
Public Sub TestInvalidFormulaReasonMatchesContext()
    CustomTestSetTitles Assert, "SetupErrors", "TestInvalidFormulaReasonMatchesContext"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetDictionaryValue hostBook, 3, "Control Details", "N(dup_variable)"
    SetDictionaryValue hostBook, 3, "Min", vbNullString
    SetDictionaryValue hostBook, 3, "Max", vbNullString

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "formula for variable ""calc_invalid"" will eventually fail"), _
                  "A formula that is invalid in the linelist context should be reported."
    Assert.IsFalse CheckingsContain(results, "The formula seems correct"), _
                   "The reason must come from the context the formula was validated in."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInvalidFormulaReasonMatchesContext", Err.Number, Err.Description
    Resume Cleanup
End Sub

'C4 has no test of its own. TidyAnalysisSheet now hands its failure reason back
'and CheckAnalysis reports it, where the old code ran the whole block under
'On Error Resume Next. Locking the Analysis sheet with a key the service does
'not hold is the only way a test can reach that failure, and a locked Analysis
'sheet makes Run refuse further down, before the report is written. Three
'shapes were tried against the live workbook: the checker still tidied the
'sheet and Run still refused later, so the assertion measured the second
'problem rather than the first one.

'@TestMethod("SetupErrors")
'Covers C5. WorksheetFunction.Trim collapses the inner double space, and the
'trimmed name was then used as the lookup key against the raw column, so the
'row answered an empty control and seven checks were skipped.
Public Sub TestSpacedVariableNameStillReadsItsControl()
    CustomTestSetTitles Assert, "SetupErrors", "TestSpacedVariableNameStillReadsItsControl"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetDictionaryValue hostBook, 5, "Variable Name", "spaced  ctrl"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "control ""unknown_control"" of variable ""spaced ctrl"" is unknown"), _
                  "A variable name holding a double space should still have its control checked."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSpacedVariableNameStillReadsItsControl", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C6. CustomTable.DataRange answers Nothing for a missing column, and the
'old code read .Column off it, so a renamed header raised error 91 out of Run
'and the user got no report at all.
Public Sub TestMissingDictionaryColumnReportsInsteadOfRaising()
    CustomTestSetTitles Assert, "SetupErrors", "TestMissingDictionaryColumnReportsInsteadOfRaising"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    RenameDictionaryHeader hostBook, "Main Label", "Renamed by hand"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "The column ""Main Label"" is missing"), _
                  "A missing dictionary column should be reported rather than raised."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingDictionaryColumnReportsInsteadOfRaising", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C7. Row 4 carries Export Number 7. The old code read the status by row
'position and the password by Export Number, so it looked for a row numbered 4,
'found none, and reported nothing at all for that export.
Public Sub TestExportChecksFollowExportNumber()
    CustomTestSetTitles Assert, "SetupErrors", "TestExportChecksFollowExportNumber"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetExportValue hostBook, 4, "export number", 7
    SetExportValue hostBook, 4, "password", vbNullString
    SetExportValue hostBook, 4, "include personal identifiers", "yes"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "The Export Number 7 has no password, but keeps identifiers"), _
                  "The export checks should name the Export Number the row carries."
    Assert.IsFalse CheckingsContain(results, "The Export Number 4 has no password"), _
                   "No finding should name an export number that has no row."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportChecksFollowExportNumber", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C8. LLExport trims every chunk before it resolves it. The old code did
'not, so " dup_variable" was reported as a variable that does not exist.
Public Sub TestFileNameChunksAreTrimmed()
    CustomTestSetTitles Assert, "SetupErrors", "TestFileNameChunksAreTrimmed"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetAllExportFileNames hostBook, SAFE_FILE_NAME
    SetExportValue hostBook, 3, "file name", """pre"" + dup_variable"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsFalse CheckingsContain(results, "which does not exists"), _
                   "A chunk with spaces around it should be trimmed before the lookup."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestFileNameChunksAreTrimmed", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C8. LLExport reads a chunk wrapped in single quotes as a literal. The
'old test looked for a double quote anywhere in the chunk, so 'my label' was
'reported as a missing variable. The literal chunk sits second in the template
'because Excel reads a leading apostrophe in a cell as the text-prefix marker
'and drops it from the stored value.
Public Sub TestSingleQuotedFileNameChunkIsLiteral()
    CustomTestSetTitles Assert, "SetupErrors", "TestSingleQuotedFileNameChunkIsLiteral"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetAllExportFileNames hostBook, SAFE_FILE_NAME
    SetExportValue hostBook, 3, "file name", """pre"" + 'my label'"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsFalse CheckingsContain(results, "which does not exists"), _
                   "A single-quoted chunk is a literal, the same way LLExport reads it."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSingleQuotedFileNameChunkIsLiteral", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C8. LLExport compares the sheet type without regard to case. The old
'code used a case-sensitive <>, so VLIST1D was reported as "not a vlist1D
'variable".
Public Sub TestUpperCaseSheetTypePassesFileNameCheck()
    CustomTestSetTitles Assert, "SetupErrors", "TestUpperCaseSheetTypePassesFileNameCheck"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetAllExportFileNames hostBook, SAFE_FILE_NAME
    SetDictionaryValue hostBook, 1, "Sheet Type", "VLIST1D"
    SetExportValue hostBook, 3, "file name", "dup_variable"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsFalse CheckingsContain(results, "which is not a vlist1D variable"), _
                   "An upper-case sheet type is a vlist1D sheet for the exporter."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpperCaseSheetTypePassesFileNameCheck", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C9. Split("", "+") gives one empty chunk, which the old code read as a
'variable named "", so an export with no file name got two findings.
Public Sub TestEmptyFileNameReportsOnlyMissingName()
    CustomTestSetTitles Assert, "SetupErrors", "TestEmptyFileNameReportsOnlyMissingName"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetAllExportFileNames hostBook, SAFE_FILE_NAME
    SetExportValue hostBook, 2, "file name", vbNullString

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "there is no value for ""File name"""), _
                  "An active export with no file name should be reported once."
    Assert.IsFalse CheckingsContain(results, "named """" which does not exists"), _
                   "An empty file name should not produce a phantom variable finding."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEmptyFileNameReportsOnlyMissingName", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C10. The Exports table stops at export number 4 while the dictionary
'counter runs to 20. Reading the status by row position walked off the end of
'the range, which reads the cell below and raises nothing, so the note fired
'for an export that does not exist.
Public Sub TestExportCounterBeyondExportRowsIsQuiet()
    CustomTestSetTitles Assert, "SetupErrors", "TestExportCounterBeyondExportRowsIsQuiet"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetDictionaryValue hostBook, 1, "Export 5", "included"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsFalse CheckingsContain(results, """Export 5"" column is filled"), _
                   "An export number with no row in the Exports table should be quiet."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportCounterBeyondExportRowsIsQuiet", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C11. CustomTable.CellRange answers Nothing for a missing column and
'IsEmpty(Nothing) is False, so a missing required column used to read as "this
'field is filled" and the check was skipped.
Public Sub TestMissingExportColumnIsReported()
    CustomTestSetTitles Assert, "SetupErrors", "TestMissingExportColumnIsReported"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    DeleteExportColumn hostBook, "password"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "The column ""Password"" is missing"), _
                  "A missing required export column should be named in the report."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingExportColumnIsReported", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C12. LLChoices.Categories matches without regard to case at build time,
'so "choice a" and "Choice A" are one category. BetterArray.Includes compares
'in binary, so the old code reported a category that exists as missing.
Public Sub TestCategoryMatchIgnoresCase()
    CustomTestSetTitles Assert, "SetupErrors", "TestCategoryMatchIgnoresCase"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SetDictionaryValue hostBook, 7, "Control Details", _
                       "CHOICE_FORMULA(LIST_primary, A1=""Yes"", ""choice a"")"

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsFalse CheckingsContain(results, "category ""choice a"" does not exists"), _
                   "A category that differs only in case should be found."
    Assert.IsFalse CheckingsContain(results, "choice formula ""LIST_primary"" is not present"), _
                   "A choice list whose name differs only in case should be found."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoryMatchIgnoresCase", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C16. Checkings used to hand out the live container, so a caller that
'cleared it emptied the checker's own results.
Public Sub TestCheckingsReturnsACopy()
    CustomTestSetTitles Assert, "SetupErrors", "TestCheckingsReturnsACopy"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    Set checker = SetupErrors.Create(hostBook)
    checker.Run

    Set results = checker.Checkings
    Assert.IsTrue results.Length > 0, "Run should produce checkings."

    results.Clear

    Assert.IsTrue checker.Checkings.Length > 0, _
                  "Clearing the returned array should leave the checker's own results alone."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCheckingsReturnsACopy", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@TestMethod("SetupErrors")
'Covers C17. With "Control Details" sitting before "Control", a partial column
'search answers the details for a request for the control, so the control
'checks read the wrong cell in silence.
Public Sub TestControlLookupIgnoresControlDetails()
    CustomTestSetTitles Assert, "SetupErrors", "TestControlLookupIgnoresControlDetails"
    On Error GoTo Fail

    Dim hostBook As Workbook
    Dim checker As SetupErrors
    Dim results As BetterArray

    Set hostBook = PrepareSetupWorkbook(includeIssues:=True)
    SwapControlColumns hostBook

    Set checker = SetupErrors.Create(hostBook)
    checker.Run
    Set results = checker.Checkings

    Assert.IsTrue CheckingsContain(results, "control ""unknown_control"" of variable ""unknown_ctrl"" is unknown"), _
                  "The control must be read from the Control column whatever the column order is."

Cleanup:
    Set checker = Nothing
    DeleteWorkbook hostBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestControlLookupIgnoresControlDetails", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@section Helpers
'===============================================================================

Private Function PrepareSetupWorkbook(Optional ByVal includeIssues As Boolean = False) As Workbook
    Dim wb As Workbook

    Set wb = NewWorkbook

    FormulaTestFixture.PrepareFormulaFixtureSheet FORMULA_SHEET, outwb:=wb
    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, wb
    SetupImportTestFixture.PrepareSetupDictionarySheet DICTIONARY_SHEET, "dup_variable", "FixtureSheet", 5, 1, wb
    SetupImportTestFixture.PrepareSetupChoicesSheet CHOICES_SHEET, 4, 1, wb
    SetupImportTestFixture.PrepareSetupExportsSheet EXPORTS_SHEET, "active", """" & "export" & """", "Export Label", 4, 1, wb
    SetupImportTestFixture.PrepareSetupAnalysisSheet ANALYSIS_SHEET, "fixture", "Analysis header", wb
    SetupImportTestFixture.PrepareSetupTranslationsSheet TRANSLATIONS_SHEET, "Tab_Translations", "Label", "Translation", "tag", 5, 2, True, wb
    EnsureWorksheet DROPDOWN_SHEET, wb, True
    EnsureWorksheet CHECK_OUTPUT_SHEET, wb, True
    EnsureWorksheet REGISTRY_SHEET, wb, True

    ConfigureDictionarySheet wb, includeIssues
    ConfigureChoicesSheet wb, includeIssues
    ConfigureExportsSheet wb, includeIssues
    ConfigureAnalysisSheet wb, includeIssues

    Set PrepareSetupWorkbook = wb
End Function

Private Sub AssertContainsMessage(ByVal results As BetterArray, ByVal expectedText As String, ByVal context As String)
    Assert.IsTrue CheckingsContain(results, expectedText), context & " should include """ & expectedText & """"
End Sub

'@description Column index of one dictionary header, matched whole-string.
Private Function DictionaryColumn(ByVal hostBook As Workbook, ByVal headerName As String) As Long
    Dim dictTable As DataSheet

    Set dictTable = DataSheet.Create(hostBook.Worksheets(DICTIONARY_SHEET), DICT_HEADER_ROW, 1)
    DictionaryColumn = dictTable.ColumnIndex(headerName, shouldExist:=True, matchCase:=False)
End Function

'@description Write one dictionary cell. dataRow 1 is the first data row.
Private Sub SetDictionaryValue(ByVal hostBook As Workbook, ByVal dataRow As Long, _
                               ByVal headerName As String, ByVal newValue As Variant)
    hostBook.Worksheets(DICTIONARY_SHEET) _
            .Cells(DICT_HEADER_ROW + dataRow, DictionaryColumn(hostBook, headerName)).Value = newValue
End Sub

'@description Rename one dictionary header so the column stops answering.
Private Sub RenameDictionaryHeader(ByVal hostBook As Workbook, ByVal headerName As String, _
                                   ByVal newHeader As String)
    hostBook.Worksheets(DICTIONARY_SHEET) _
            .Cells(DICT_HEADER_ROW, DictionaryColumn(hostBook, headerName)).Value = newHeader
End Sub

'@description Put Control Details ahead of Control, values and headers together.
Private Sub SwapControlColumns(ByVal hostBook As Workbook)
    Dim sh As Worksheet
    Dim controlCol As Long
    Dim detailsCol As Long
    Dim rowIdx As Long
    Dim controlText As String
    Dim detailsText As String

    Set sh = hostBook.Worksheets(DICTIONARY_SHEET)
    controlCol = DictionaryColumn(hostBook, "Control")
    detailsCol = DictionaryColumn(hostBook, "Control Details")

    For rowIdx = DICT_HEADER_ROW + 1 To DICT_LAST_DATA_ROW
        controlText = CStr(sh.Cells(rowIdx, controlCol).Value)
        detailsText = CStr(sh.Cells(rowIdx, detailsCol).Value)
        sh.Cells(rowIdx, controlCol).Value = detailsText
        sh.Cells(rowIdx, detailsCol).Value = controlText
    Next rowIdx

    sh.Cells(DICT_HEADER_ROW, controlCol).Value = "Control Details"
    sh.Cells(DICT_HEADER_ROW, detailsCol).Value = "Control"
End Sub

'@description Write one cell of the exports table by column name and row index.
Private Sub SetExportValue(ByVal hostBook As Workbook, ByVal rowIndex As Long, _
                           ByVal columnName As String, ByVal newValue As Variant)
    hostBook.Worksheets(EXPORTS_SHEET).ListObjects(1) _
            .ListColumns(columnName).DataBodyRange.Cells(rowIndex, 1).Value = newValue
End Sub

'@description Give every export the same file name, so one test can change one.
Private Sub SetAllExportFileNames(ByVal hostBook As Workbook, ByVal newValue As String)
    Dim lo As ListObject
    Dim idx As Long

    Set lo = hostBook.Worksheets(EXPORTS_SHEET).ListObjects(1)

    For idx = 1 To lo.ListRows.Count
        lo.ListColumns("file name").DataBodyRange.Cells(idx, 1).Value = newValue
    Next idx
End Sub

'@description Drop one column of the exports table.
Private Sub DeleteExportColumn(ByVal hostBook As Workbook, ByVal columnName As String)
    hostBook.Worksheets(EXPORTS_SHEET).ListObjects(1).ListColumns(columnName).Delete
End Sub

'@description Answer whether a worksheet holds a piece of text anywhere.
Private Function SheetHoldsText(ByVal sh As Worksheet, ByVal needle As String) As Boolean
    Dim found As Range

    On Error Resume Next
        Set found = sh.UsedRange.Find(What:=needle, LookIn:=xlValues, LookAt:=xlPart)
    On Error GoTo 0

    SheetHoldsText = Not (found Is Nothing)
End Function

Private Sub ConfigureDictionarySheet(ByVal hostBook As Workbook, ByVal includeIssues As Boolean)
    Dim dictSheet As Worksheet
    Dim dictTable As DataSheet
    Dim controlCol As Long
    Dim detailsCol As Long
    Dim sheetCol As Long
    Dim mainCol As Long
    Dim varCol As Long
    Dim uniqueCol As Long
    Dim minCol As Long
    Dim maxCol As Long
    Dim varTypeCol As Long
    Dim formatCol As Long
    Dim sheetTypeCol As Long
    Dim export1Col As Long
    Dim export2Col As Long
    Dim export4Col As Long
    Dim statusCol As Long
    Dim firstDataRow As Long
    Dim secondDataRow As Long
    Dim thirdDataRow As Long
    Dim fourthDataRow As Long
    Dim fifthDataRow As Long
    Dim sixthDataRow As Long
    Dim seventhDataRow As Long
    Dim lastDataRow As Long
    Dim totalColumns As Long
    Dim lastColumn As Long
    Dim dataStart As Range
    Dim dataArea As Range
    Dim tableRange As Range
    Dim targetTable As ListObject

    Set dictSheet = hostBook.Worksheets(DICTIONARY_SHEET)
    Set dictTable = DataSheet.Create(dictSheet, 5, 1)

    varCol = dictTable.ColumnIndex("Variable Name", shouldExist:=True, matchCase:=False)
    mainCol = dictTable.ColumnIndex("Main Label", shouldExist:=True, matchCase:=False)
    sheetCol = dictTable.ColumnIndex("Sheet Name", shouldExist:=True, matchCase:=False)
    controlCol = dictTable.ColumnIndex("Control", shouldExist:=True, matchCase:=False)
    detailsCol = dictTable.ColumnIndex("Control Details", shouldExist:=True, matchCase:=False)
    uniqueCol = dictTable.ColumnIndex("Unique", shouldExist:=True, matchCase:=False)
    minCol = dictTable.ColumnIndex("Min", shouldExist:=True, matchCase:=False)
    maxCol = dictTable.ColumnIndex("Max", shouldExist:=True, matchCase:=False)
    varTypeCol = dictTable.ColumnIndex("Variable Type", shouldExist:=True, matchCase:=False)
    formatCol = dictTable.ColumnIndex("Variable Format", shouldExist:=True, matchCase:=False)
    sheetTypeCol = dictTable.ColumnIndex("Sheet Type", shouldExist:=True, matchCase:=False)
    export1Col = dictTable.ColumnIndex("Export 1", shouldExist:=True, matchCase:=False)
    export2Col = dictTable.ColumnIndex("Export 2", shouldExist:=True, matchCase:=False)
    export4Col = dictTable.ColumnIndex("Export 4", shouldExist:=True, matchCase:=False)
    statusCol = dictTable.ColumnIndex("Status", shouldExist:=True, matchCase:=False)

    firstDataRow = dictTable.DataStartRow + 1
    secondDataRow = firstDataRow + 1
    thirdDataRow = secondDataRow + 1
    fourthDataRow = thirdDataRow + 1
    fifthDataRow = fourthDataRow + 1
    sixthDataRow = fifthDataRow + 1
    seventhDataRow = sixthDataRow + 1
    lastDataRow = IIf(includeIssues, seventhDataRow, secondDataRow)

    lastColumn = dictSheet.Cells(dictTable.DataStartRow, dictSheet.Columns.Count).End(xlToLeft).Column
    totalColumns = lastColumn - dictTable.DataStartColumn + 1

    Set dataStart = dictSheet.Cells(firstDataRow, dictTable.DataStartColumn)
    Set dataArea = dictSheet.Range(dataStart, dataStart.Offset((lastDataRow - firstDataRow) + 5, totalColumns - 1))
    dataArea.ClearContents

    dictSheet.Cells(firstDataRow, varCol).Value = "dup_variable"
    dictSheet.Cells(firstDataRow, mainCol).Value = "Duplicate variable label"
    dictSheet.Cells(firstDataRow, sheetCol).Value = "FixtureSheet"
    dictSheet.Cells(firstDataRow, controlCol).Value = "choice_manual"
    dictSheet.Cells(firstDataRow, detailsCol).Value = "missing_choice"
    dictSheet.Cells(firstDataRow, uniqueCol).Value = "yes"
    dictSheet.Cells(firstDataRow, varTypeCol).Value = "text"
    dictSheet.Cells(firstDataRow, sheetTypeCol).Value = "vlist1D"
    dictSheet.Cells(firstDataRow, statusCol).Value = "active"
    dictSheet.Cells(firstDataRow, export1Col).Value = "included"

    dictSheet.Cells(secondDataRow, varCol).Value = "dup_variable"
    dictSheet.Cells(secondDataRow, mainCol).Value = vbNullString
    dictSheet.Cells(secondDataRow, sheetCol).Value = vbNullString
    dictSheet.Cells(secondDataRow, controlCol).Value = "choice_manual"
    dictSheet.Cells(secondDataRow, detailsCol).Value = "missing_choice"
    dictSheet.Cells(secondDataRow, varTypeCol).Value = "text"
    dictSheet.Cells(secondDataRow, sheetTypeCol).Value = "vlist1D"
    dictSheet.Cells(secondDataRow, statusCol).Value = "inactive"

    If includeIssues Then
        dictSheet.Cells(thirdDataRow, varCol).Value = "calc_invalid"
        dictSheet.Cells(thirdDataRow, mainCol).Value = "Calc Invalid"
        dictSheet.Cells(thirdDataRow, sheetCol).Value = "CalcSheet"
        dictSheet.Cells(thirdDataRow, controlCol).Value = "formula"
        dictSheet.Cells(thirdDataRow, detailsCol).Value = "INVALID("
        dictSheet.Cells(thirdDataRow, minCol).Value = "INVALID("
        dictSheet.Cells(thirdDataRow, maxCol).Value = "INVALID("
        dictSheet.Cells(thirdDataRow, varTypeCol).Value = vbNullString
        dictSheet.Cells(thirdDataRow, formatCol).Value = "General"
        dictSheet.Cells(thirdDataRow, sheetTypeCol).Value = "hlist2D"
        dictSheet.Cells(thirdDataRow, statusCol).Value = "active"

        dictSheet.Cells(fourthDataRow, varCol).Value = "abc"
        dictSheet.Cells(fourthDataRow, mainCol).Value = "Short variable"
        dictSheet.Cells(fourthDataRow, sheetCol).Value = "ShortSheet"
        dictSheet.Cells(fourthDataRow, controlCol).Value = "text"
        dictSheet.Cells(fourthDataRow, varTypeCol).Value = "text"
        dictSheet.Cells(fourthDataRow, sheetTypeCol).Value = "vlist1D"
        dictSheet.Cells(fourthDataRow, statusCol).Value = "active"

        dictSheet.Cells(fifthDataRow, varCol).Value = "unknown_ctrl"
        dictSheet.Cells(fifthDataRow, mainCol).Value = "Unknown control"
        dictSheet.Cells(fifthDataRow, sheetCol).Value = "UnknownSheet"
        dictSheet.Cells(fifthDataRow, controlCol).Value = "unknown_control"
        dictSheet.Cells(fifthDataRow, varTypeCol).Value = "text"
        dictSheet.Cells(fifthDataRow, sheetTypeCol).Value = "vlist1D"
        dictSheet.Cells(fifthDataRow, statusCol).Value = "active"

        dictSheet.Cells(sixthDataRow, varCol).Value = "missing_choice_formula"
        dictSheet.Cells(sixthDataRow, mainCol).Value = "Missing choice formula"
        dictSheet.Cells(sixthDataRow, sheetCol).Value = "FormulaSheet"
        dictSheet.Cells(sixthDataRow, controlCol).Value = "choice_formula"
        dictSheet.Cells(sixthDataRow, detailsCol).Value = "CHOICE_FORMULA(missing_choice, A1=""Yes"", ""Missing formula"")"
        dictSheet.Cells(sixthDataRow, varTypeCol).Value = "text"
        dictSheet.Cells(sixthDataRow, sheetTypeCol).Value = "vlist1D"
        dictSheet.Cells(sixthDataRow, statusCol).Value = "active"

        dictSheet.Cells(seventhDataRow, varCol).Value = "choice_bad_category"
        dictSheet.Cells(seventhDataRow, mainCol).Value = "Choice bad category"
        dictSheet.Cells(seventhDataRow, sheetCol).Value = "FormulaSheet"
        dictSheet.Cells(seventhDataRow, controlCol).Value = "choice_formula"
        dictSheet.Cells(seventhDataRow, detailsCol).Value = "CHOICE_FORMULA(list_primary, A1=""Yes"", ""Unknown option"")"
        dictSheet.Cells(seventhDataRow, varTypeCol).Value = "text"
        dictSheet.Cells(seventhDataRow, sheetTypeCol).Value = "vlist1D"
        dictSheet.Cells(seventhDataRow, statusCol).Value = "active"

        dictSheet.Cells(dictTable.DataStartRow, export4Col).Value = vbNullString
    End If

    Set tableRange = dictSheet.Range(dictSheet.Cells(dictTable.DataStartRow, dictTable.DataStartColumn), _
                                     dictSheet.Cells(lastDataRow, dictTable.DataStartColumn + totalColumns - 1))
    If dictSheet.ListObjects.Count = 0 Then
        Set targetTable = dictSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
        targetTable.Name = "TST_Dictionary"
    Else
        dictSheet.ListObjects(1).Resize tableRange
    End If
End Sub

Private Sub ConfigureChoicesSheet(ByVal hostBook As Workbook, ByVal includeIssues As Boolean)
    Dim choicesSheet As Worksheet
    Dim choicesTable As DataSheet
    Dim listNameCol As Long
    Dim orderCol As Long
    Dim labelCol As Long
    Dim shortCol As Long
    Dim firstDataRow As Long
    Dim secondDataRow As Long
    Dim thirdDataRow As Long
    Dim fourthDataRow As Long
    Dim fifthDataRow As Long
    Dim sixthDataRow As Long
    Dim lastDataRow As Long
    Dim totalColumns As Long
    Dim lastColumn As Long
    Dim dataRange As Range
    Dim tableRange As Range
    Dim targetTable As ListObject

    Set choicesSheet = hostBook.Worksheets(CHOICES_SHEET)
    Set choicesTable = DataSheet.Create(choicesSheet, 4, 1)

    listNameCol = choicesTable.ColumnIndex("list name", shouldExist:=True, matchCase:=False)
    orderCol = choicesTable.ColumnIndex("ordering list", shouldExist:=True, matchCase:=False)
    labelCol = choicesTable.ColumnIndex("label", shouldExist:=True, matchCase:=False)
    shortCol = choicesTable.ColumnIndex("short label", shouldExist:=True, matchCase:=False)

    firstDataRow = choicesTable.DataStartRow + 1
    secondDataRow = firstDataRow + 1
    thirdDataRow = secondDataRow + 1
    fourthDataRow = thirdDataRow + 1
    fifthDataRow = fourthDataRow + 1
    sixthDataRow = fifthDataRow + 1
    lastDataRow = IIf(includeIssues, sixthDataRow, secondDataRow)

    lastColumn = choicesSheet.Cells(choicesTable.DataStartRow, choicesSheet.Columns.Count).End(xlToLeft).Column
    totalColumns = lastColumn - choicesTable.DataStartColumn + 1

    Set dataRange = choicesSheet.Range(choicesSheet.Cells(firstDataRow, choicesTable.DataStartColumn), _
                                       choicesSheet.Cells(lastDataRow + 3, choicesTable.DataStartColumn + totalColumns - 1))
    dataRange.ClearContents

    choicesSheet.Cells(firstDataRow, listNameCol).Value = "list_primary"
    choicesSheet.Cells(firstDataRow, orderCol).Value = 1
    choicesSheet.Cells(firstDataRow, labelCol).Value = "Choice A"
    choicesSheet.Cells(firstDataRow, shortCol).Value = "A"

    If includeIssues Then
        choicesSheet.Cells(secondDataRow, listNameCol).Value = "list_primary"
        choicesSheet.Cells(secondDataRow, orderCol).Value = 2
        choicesSheet.Cells(secondDataRow, labelCol).Value = vbNullString
        choicesSheet.Cells(secondDataRow, shortCol).Value = "B"

        choicesSheet.Cells(thirdDataRow, listNameCol).Value = "list_secondary"
        choicesSheet.Cells(thirdDataRow, orderCol).Value = 1
        choicesSheet.Cells(thirdDataRow, labelCol).Value = "Option 1"
        choicesSheet.Cells(thirdDataRow, shortCol).Value = "Opt1"

        choicesSheet.Cells(fourthDataRow, listNameCol).Value = vbNullString
        choicesSheet.Cells(fourthDataRow, orderCol).Value = vbNullString
        choicesSheet.Cells(fourthDataRow, labelCol).Value = "Orphan label"
        choicesSheet.Cells(fourthDataRow, shortCol).Value = "Orphan"

        choicesSheet.Cells(fifthDataRow, listNameCol).Value = vbNullString
        choicesSheet.Cells(fifthDataRow, orderCol).Value = 5
        choicesSheet.Cells(fifthDataRow, labelCol).Value = "Ordered orphan"
        choicesSheet.Cells(fifthDataRow, shortCol).Value = "Ordered"

        choicesSheet.Cells(sixthDataRow, listNameCol).Value = "list_missing_order"
        choicesSheet.Cells(sixthDataRow, orderCol).Value = vbNullString
        choicesSheet.Cells(sixthDataRow, labelCol).Value = "Missing order"
        choicesSheet.Cells(sixthDataRow, shortCol).Value = "Missing"
    End If

    Set tableRange = choicesSheet.Range(choicesSheet.Cells(choicesTable.DataStartRow, choicesTable.DataStartColumn), _
                                    choicesSheet.Cells(lastDataRow, choicesTable.DataStartColumn + totalColumns - 1))
    If choicesSheet.ListObjects.Count = 0 Then
        Set targetTable = choicesSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
        targetTable.Name = "TST_Choices"
    Else
        choicesSheet.ListObjects(1).Resize tableRange
    End If
End Sub

Private Sub ConfigureExportsSheet(ByVal hostBook As Workbook, ByVal includeIssues As Boolean)
    Dim exportsSheet As Worksheet
    Dim exportsTable As ListObject
    Dim dataBody As Range
    Dim maxRows As Long
    Dim rowIndex As Long

    Set exportsSheet = hostBook.Worksheets(EXPORTS_SHEET)
    Set exportsTable = exportsSheet.ListObjects(1)

    Set dataBody = exportsTable.DataBodyRange
    If Not dataBody Is Nothing Then dataBody.ClearContents

    maxRows = IIf(includeIssues, 4, 1)

    Do While exportsTable.ListRows.Count < maxRows
        exportsTable.ListRows.Add
    Loop

    Do While exportsTable.ListRows.Count > maxRows
        exportsTable.ListRows(exportsTable.ListRows.Count).Delete
    Loop

    For rowIndex = 1 To exportsTable.ListRows.Count
        exportsTable.ListColumns("export number").DataBodyRange.Cells(rowIndex, 1).Value = rowIndex
        exportsTable.ListColumns("status").DataBodyRange.Cells(rowIndex, 1).Value = IIf(includeIssues And rowIndex = 1, "inactive", "active")
        exportsTable.ListColumns("label button").DataBodyRange.Cells(rowIndex, 1).Value = "Export " & rowIndex
        exportsTable.ListColumns("file format").DataBodyRange.Cells(rowIndex, 1).Value = "xlsx"
        exportsTable.ListColumns("file name").DataBodyRange.Cells(rowIndex, 1).Value = "export" & rowIndex & ".xlsx"
        exportsTable.ListColumns("password").DataBodyRange.Cells(rowIndex, 1).Value = "yes"
        exportsTable.ListColumns("include personal identifiers").DataBodyRange.Cells(rowIndex, 1).Value = "no"
        exportsTable.ListColumns("header format").DataBodyRange.Cells(rowIndex, 1).Value = "variables names"
        exportsTable.ListColumns("export metadata sheets").DataBodyRange.Cells(rowIndex, 1).Value = "yes"
        exportsTable.ListColumns("export analyses sheets").DataBodyRange.Cells(rowIndex, 1).Value = "no"
    Next rowIndex

    If includeIssues Then
        exportsTable.ListColumns("status").DataBodyRange.Cells(2, 1).Value = "active"
        exportsTable.ListColumns("label button").DataBodyRange.Cells(2, 1).Value = vbNullString
        exportsTable.ListColumns("file format").DataBodyRange.Cells(2, 1).Value = vbNullString
        exportsTable.ListColumns("file name").DataBodyRange.Cells(2, 1).Value = vbNullString
        exportsTable.ListColumns("password").DataBodyRange.Cells(2, 1).Value = vbNullString
        exportsTable.ListColumns("include personal identifiers").DataBodyRange.Cells(2, 1).Value = "yes"
        exportsTable.ListColumns("header format").DataBodyRange.Cells(2, 1).Value = vbNullString
        exportsTable.ListColumns("export metadata sheets").DataBodyRange.Cells(2, 1).Value = vbNullString

        exportsTable.ListColumns("status").DataBodyRange.Cells(3, 1).Value = "active"
        exportsTable.ListColumns("label button").DataBodyRange.Cells(3, 1).Value = "Valid export"
        exportsTable.ListColumns("file format").DataBodyRange.Cells(3, 1).Value = "xlsx"
        exportsTable.ListColumns("file name").DataBodyRange.Cells(3, 1).Value = "missing_var+calc_invalid"
        exportsTable.ListColumns("password").DataBodyRange.Cells(3, 1).Value = "yes"
        exportsTable.ListColumns("include personal identifiers").DataBodyRange.Cells(3, 1).Value = "no"
        exportsTable.ListColumns("header format").DataBodyRange.Cells(3, 1).Value = "variables names"

        exportsTable.ListColumns("status").DataBodyRange.Cells(4, 1).Value = "active"
        exportsTable.ListColumns("label button").DataBodyRange.Cells(4, 1).Value = "Missing dictionary column"
        exportsTable.ListColumns("file format").DataBodyRange.Cells(4, 1).Value = "xlsx"
        exportsTable.ListColumns("file name").DataBodyRange.Cells(4, 1).Value = """literal"""
        exportsTable.ListColumns("password").DataBodyRange.Cells(4, 1).Value = "yes"
        exportsTable.ListColumns("include personal identifiers").DataBodyRange.Cells(4, 1).Value = "no"
        exportsTable.ListColumns("header format").DataBodyRange.Cells(4, 1).Value = "variables names"
    End If
End Sub

Private Sub AddEmptyAnalysisRow(ByVal hostBook As Workbook)
    Dim analysisSheet As Worksheet
    Dim lo As ListObject

    Set analysisSheet = hostBook.Worksheets(ANALYSIS_SHEET)
    Set lo = analysisSheet.ListObjects("Tab_Univariate_Analysis")
    'Insert inside the data body (not at the boundary) to avoid
    'stacked-table conflicts. The ListObject auto-expands and
    'the new row is empty. Existing data shifts to position 2.
    analysisSheet.Rows(lo.DataBodyRange.Row).Insert Shift:=xlShiftDown
End Sub


Private Sub ConfigureAnalysisSheet(ByVal hostBook As Workbook, ByVal includeIssues As Boolean)
    Dim analysisSheet As Worksheet
    Dim lo As ListObject
    Dim typeLabel As String

    Set analysisSheet = hostBook.Worksheets(ANALYSIS_SHEET)

    For Each lo In analysisSheet.ListObjects
        typeLabel = AnalysisTableTypeLabel(lo.Name)
        If LenB(typeLabel) > 0 Then
            analysisSheet.Cells(lo.HeaderRowRange.Row - 2, 1).Value = typeLabel
        End If
    Next lo
End Sub

Private Function AnalysisTableTypeLabel(ByVal tableName As String) As String
    Select Case tableName
    Case "Tab_Global_Summary"
        AnalysisTableTypeLabel = "Global summary"
    Case "Tab_Univariate_Analysis"
        AnalysisTableTypeLabel = "Univariate analysis"
    Case "Tab_Bivariate_Analysis"
        AnalysisTableTypeLabel = "Bivariate analysis"
    Case "Tab_TimeSeries_Analysis"
        AnalysisTableTypeLabel = "Time series analysis"
    Case "Tab_Spatial_Analysis"
        AnalysisTableTypeLabel = "Spatial analysis"
    Case "Tab_Graph_TimeSeries"
        AnalysisTableTypeLabel = "Time series graphs"
    Case "Tab_SpatioTemporal_Analysis"
        AnalysisTableTypeLabel = "Spatio-temporal analysis"
    Case Else
        AnalysisTableTypeLabel = vbNullString
    End Select
End Function

Private Function CheckingsContain(ByVal source As BetterArray, ByVal expectedText As String) As Boolean
    Dim idx As Long
    Dim keyIdx As Long
    Dim check As Checking
    Dim keys As BetterArray
    Dim keyName As String
    Dim labelValue As String

    If source Is Nothing Then Exit Function
    If source.Length = 0 Then Exit Function

    For idx = source.LowerBound To source.UpperBound
        Set check = source.Item(idx)
        If Not check Is Nothing Then
            Set keys = check.ListOfKeys
            If Not keys Is Nothing Then
                For keyIdx = keys.LowerBound To keys.UpperBound
                    keyName = CStr(keys.Item(keyIdx))
                    labelValue = check.ValueOf(keyName)
                    If InStr(1, labelValue, expectedText, vbTextCompare) > 0 Then
                        CheckingsContain = True
                        Exit Function
                    End If
                Next keyIdx
            End If
        End If
    Next idx
End Function
