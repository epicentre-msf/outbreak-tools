Attribute VB_Name = "TestHeadlessSetupFill"
Attribute VB_Description = "Fills a real setup workbook from the generic setup, headless"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Fills a real setup workbook from the generic setup, headless")

Option Explicit

'@description
'Step one of the headless workflow, under test: a real setup workbook is
'filled from another setup workbook with no ribbon, no form and no dialog.
'
'WHAT IT PROVES
'-------------------------------------------------------------------------------
'That HeadlessBuild.ImportSetupFromWorkbook can inject its import module into a
'target setup, run the production import sequence inside it, and leave a
'workbook whose dictionary carries the source's rows. The filled file is left
'on disk, because step two of the workflow builds a linelist from it.
'
'THE TWO FILES IT NEEDS
'-------------------------------------------------------------------------------
'  src/bin/setup/setup_dev.xlsb          the setup to fill, copied per run
'  src/bin/test-files/generic-test-setup.xlsb   the filled setup read from
'
'Both are gitignored binaries carried by the asset store, the same as every
'other workbook this project builds against. A missing file is reported as a
'failure naming the path rather than a raise, so a checkout without the assets
'says what to pull rather than dying in a dialog.
'
'THE PATHS ARE ABSOLUTE, AND THAT IS A KNOWN LIMIT
'-------------------------------------------------------------------------------
'A test module runs from the staging copy of the driver workbook, so a
'relative path resolves against the run directory rather than the repository.
'REPO_ROOT is the one place that assumption lives; HeadlessBuild itself takes
'every path as a parameter and holds none.
'@depends HeadlessBuild, CustomTest

Private Assert As CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const REPO_ROOT As String = _
    "/Users/komlaviamevoin/Unsync-Working-Folders/outbreak-tools"

Private Const EMPTY_SETUP As String = REPO_ROOT & "/src/bin/setup/setup_dev.xlsb"
Private Const GENERIC_SETUP As String = REPO_ROOT & "/src/bin/test-files/generic-test-setup.xlsb"
Private Const INJECTED_MODULE As String = REPO_ROOT & "/scripts/headless/vba/OBTSetupImportHeadless.bas"
Private Const FILLED_SETUP As String = REPO_ROOT & "/.obt/draft/demo_setup_filled.xlsb"

Private Const DICTIONARY_SHEET As String = "Dictionary"

Private Outcome As String
Private DictionaryRows As Long
Private SetupError As Long
Private SetupMessage As String


'@section Lifecycle
'===============================================================================

'@sub-title Fill the setup once, and read what landed.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestHeadlessSetupFill"

    SetupError = 0
    SetupMessage = vbNullString
    DictionaryRows = -1

    'One grant for everything this suite touches, before its first Dir$. An
    'already-granted machine sees no dialog.
    HeadlessBuild.EnsureFileAccess Array(REPO_ROOT)

    On Error Resume Next
        FillTheSetup
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Hand the screen back and print.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    'The fill opens and closes a workbook, so the screen goes back to the
    'driver before PrintResults writes into one of its own sheets.
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section The fill
'===============================================================================

'@sub-title Copy the empty setup, fill it from the generic one, count what landed.
Private Sub FillTheSetup()
    If Len(Dir$(EMPTY_SETUP)) = 0 Then
        Outcome = "MISSING: " & EMPTY_SETUP
        Exit Sub
    End If

    If Len(Dir$(GENERIC_SETUP)) = 0 Then
        Outcome = "MISSING: " & GENERIC_SETUP
        Exit Sub
    End If

    On Error Resume Next
        Kill FILLED_SETUP
    On Error GoTo 0

    FileCopy EMPTY_SETUP, FILLED_SETUP

    Outcome = HeadlessBuild.ImportSetupFromWorkbook(FILLED_SETUP, GENERIC_SETUP, _
                                                    INJECTED_MODULE)

    DictionaryRows = RowsOfFirstTable(FILLED_SETUP, DICTIONARY_SHEET)
End Sub

'@fun-title Reopen a workbook and count the rows of the first table of one sheet.
'@param bookPath String. The workbook to read.
'@param sheetName String. The worksheet holding the table.
'@return Long. The row count, 0 for an empty table, -1 when nothing resolves.
Private Function RowsOfFirstTable(ByVal bookPath As String, _
                                  ByVal sheetName As String) As Long
    Dim wkb As Workbook
    Dim sh As Worksheet
    Dim Lo As ListObject

    RowsOfFirstTable = -1
    If Len(Dir$(bookPath)) = 0 Then Exit Function

    On Error GoTo CleanExit
    Set wkb = Application.Workbooks.Open(fileName:=bookPath, ReadOnly:=True)

    On Error Resume Next
        Set sh = wkb.Worksheets(sheetName)
    On Error GoTo CleanExit

    If Not sh Is Nothing Then
        If sh.ListObjects.Count > 0 Then
            Set Lo = sh.ListObjects(1)
            If Lo.DataBodyRange Is Nothing Then
                RowsOfFirstTable = 0
            Else
                RowsOfFirstTable = Lo.DataBodyRange.Rows.Count
            End If
        End If
    End If

CleanExit:
    On Error Resume Next
        If Not wkb Is Nothing Then wkb.Close SaveChanges:=False
    On Error GoTo 0
End Function


'@section What the run reports
'===============================================================================

'@sub-title The import answered OK.
'@TestMethod("headless")
Private Sub TestTheHeadlessImportAnswersOK()
    Const TESTNAME As String = "TestTheHeadlessImportAnswersOK"

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, "TestTheHeadlessImportAnswersOK", "The headless setup fill"
    Assert.AreEqual "OK", Outcome, _
                    "the injected import answers OK, and names its fault otherwise"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The filled setup carries the dictionary of the source.
'@TestMethod("headless")
Private Sub TestTheFilledSetupCarriesADictionary()
    Const TESTNAME As String = "TestTheFilledSetupCarriesADictionary"

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, "TestTheFilledSetupCarriesADictionary", "The filled setup"
    Assert.IsTrue DictionaryRows > 0, _
                  "the dictionary of the filled setup holds rows (" & _
                  CStr(DictionaryRows) & ")"
    Assert.IsTrue Len(Dir$(FILLED_SETUP)) > 0, _
                  "the filled setup is on disk for the build step to read"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The fixture reports its own failure, once.
'@TestMethod("headless")
Private Sub TestTheFillRanAtAll()
    Const TESTNAME As String = "TestTheFillRanAtAll"

    If SetupError = 0 Then
        CustomTestSetTitles Assert, "TestTheFillRanAtAll", "The fill"
        Assert.AreEqual 0&, SetupError, "the fill raised nothing"
        Exit Sub
    End If

    CustomTestLogFailure Assert, TESTNAME, SetupError, _
                         "The fill could not run - " & SetupMessage
End Sub
