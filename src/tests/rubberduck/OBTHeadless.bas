Attribute VB_Name = "OBTHeadless"
Attribute VB_Description = "Headless, AppleScript-callable test runner and testsOutputs serializer"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("Headless, AppleScript-callable test runner and testsOutputs serializer")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

' =============================================================================
' Headless entry point for the macOS AppleScript test loop.
'
' Self-contained on purpose: depends only on the test modules (invoked through
' Application.Run) and on VBProject access for '@TestMethod discovery. It does
' NOT use TestHelpers (BusyApp/EnsureWorksheet) or the ribbon callback
' clickRibbonTests, so the workbook needs neither TestHelpers nor
' CustomTestImplementation to run headless.
'
'   OBTRunAllTests  -- parameterless, `run VB macro`-callable. Suppresses UI,
'                      runs every module listed in the `ModulesForTesting`
'                      Name on `Codes`, serializes `testsOutputs` to
'                      test-results.csv next to the workbook, then Saves.
'                      No dialog, no Quit (Quit is AppleScript's job).
'
' Sub names carry no underscore on purpose: in a class/document module
' `Foo_Bar` is parsed as event `Bar` of object `Foo`, so the convention is
' kept everywhere to stay unambiguous.
' =============================================================================

Private Const CODESHEET         As String = "Codes"
Private Const MODULES_NAME      As String = "ModulesForTesting"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const RESULTS_FILENAME  As String = "test-results.csv"

' testsOutputs layout written by CheckingOutput (row 4 down):
'   col B = test title, col C = "<symbol> Success|Error", col D = label, col E = message
Private Const COL_TITLE   As Long = 2
Private Const COL_TYPE    As Long = 3
Private Const COL_LABEL   As Long = 4
Private Const COL_MESSAGE As Long = 5
Private Const FIRST_DATA_ROW As Long = 4

' --- run diagnostics (surfaced in the CSV summary line) ----------------------
Private mModulesRun As Long
Private mTestsFound As Long
Private mVbeAccess As String
Private mNameStatus As String

' True only while OBTRunAllTests is executing. CustomTest.PrintResults reads this
' (late-bound via OBTHeadlessActive) to skip the code-module injection that
' stalls under Apple Events automation.
Private mHeadlessActive As Boolean

'@sub-title Report whether a headless run is in progress (read by CustomTest.PrintResults).
Public Function OBTHeadlessActive() As Boolean
    OBTHeadlessActive = mHeadlessActive
End Function

'@EntryPoint
'@sub-title Run every registered module, serialize results, save. Dialog-free.
Public Sub OBTRunAllTests()
    Dim prevScreen As Boolean
    Dim prevAlerts As Boolean
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation

    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanExit

    mModulesRun = 0
    mTestsFound = 0
    mVbeAccess = "unknown"
    mNameStatus = "unknown"
    mHeadlessActive = True

    ResetOutputSheet
    RunAllTestModules
    SerializeTestOutputs ResultsPath()
    ThisWorkbook.Save

CleanExit:
    mHeadlessActive = False
    If Err.Number <> 0 Then
        Debug.Print "OBTRunAllTests error " & Err.Number & ": " & Err.Description
        LogRunStep "OBTRunAllTests error " & Err.Number & ": " & Err.Description
    End If

    'The restore runs inside the handler, where a second error is not
    'trappable: it kills the macro, and Application.Run then answers the
    'caller with a bare -50 on a run whose results are already written and
    'saved. Assigning Calculation can flush pending calculations and raise
    '1004, so every step here is scoped.
    On Error Resume Next
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        If Err.Number <> 0 Then
            LogRunStep "OBTRunAllTests restore error " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
    On Error GoTo 0
End Sub

'@sub-title Append one line to the diagnostics log next to the workbook.
'@details
'The same file OBTImport writes, opened and closed per line so the last line
'always survives. Failures here are swallowed: a log write must never be what
'stops a run.
'@param message String. The line to write.
Private Sub LogRunStep(ByVal message As String)
    Dim fileNum As Integer

    On Error Resume Next
        fileNum = FreeFile
        Open ThisWorkbook.Path & Application.PathSeparator & "obt-import.log" For Append As #fileNum
        Print #fileNum, message
        Close #fileNum
    On Error GoTo 0
End Sub

'@sub-title Clear testsOutputs and its CheckingOutput helper Names so the run starts clean.
'@details Sheet deletion is unreliable on macOS Excel, so hard-clear the contents,
' tables, and both worksheet- and workbook-scoped helper names (notably the
' CheckingOutput row marker) instead. This keeps re-runs idempotent - otherwise
' PrintResults appends beneath whatever a previous run left behind.
Private Sub ResetOutputSheet()
    Dim sh As Worksheet
    Dim nm As Name

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(TEST_OUTPUT_SHEET)
    On Error GoTo 0

    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
            ' Drop tables, then worksheet-scoped names (incl. the row marker), then cells.
            Do While sh.ListObjects.Count > 0
                sh.ListObjects(1).Delete
            Loop
            For Each nm In sh.Names
                nm.Delete
            Next nm
            sh.Cells.Clear
        On Error GoTo 0
    End If

    ' Remove any lingering workbook-scoped helper names from a previous run.
    On Error Resume Next
        For Each nm In ThisWorkbook.Names
            If InStr(1, nm.Name, "CheckingOutput", vbTextCompare) > 0 _
               Or InStr(1, nm.RefersTo, TEST_OUTPUT_SHEET, vbTextCompare) > 0 Then
                nm.Delete
            End If
        Next nm
    On Error GoTo 0
End Sub

'@sub-title Iterate the ModulesForTesting ListObject on Codes and run each module.
Private Sub RunAllTestModules()
    Dim rng As Range
    Dim cel As Range
    Dim moduleName As String

    Set rng = ResolveModulesRange()
    If rng Is Nothing Then Exit Sub

    For Each cel In rng.Cells
        moduleName = Trim$(CStr(cel.Value))
        If LenB(moduleName) > 0 Then
            mModulesRun = mModulesRun + 1
            RunModule moduleName
        End If
    Next cel
End Sub

'@fun-title Resolve the module-name range from the ModulesForTesting ListObject.
'@details ModulesForTesting is a ListObject (table) on Codes, not a defined Name,
' so read its first data column. A header-cell scan is kept as a resilience
' fallback in case the table name ever drifts from its header text.
Private Function ResolveModulesRange() As Range
    Dim lo As ListObject
    Dim rng As Range

    On Error Resume Next
        Set lo = ThisWorkbook.Worksheets(CODESHEET).ListObjects(MODULES_NAME)
    On Error GoTo 0

    If Not lo Is Nothing Then
        On Error Resume Next
            Set rng = lo.ListColumns(1).DataBodyRange
        On Error GoTo 0
        If Not rng Is Nothing Then
            mNameStatus = "listobject:" & rng.Address(False, False)
            Set ResolveModulesRange = rng
            Exit Function
        End If
    End If

    ' Fallback: anchor on the "ModulesForTesting" header cell and read the run of
    ' module names beneath it (cell reads always work on macOS Excel).
    Set rng = FindByHeader()
    If Not rng Is Nothing Then
        mNameStatus = "header:" & rng.Address(False, False)
        Set ResolveModulesRange = rng
        Exit Function
    End If

    mNameStatus = "not-found"
End Function

'@fun-title Locate the ModulesForTesting header cell on Codes and return the run below it.
Private Function FindByHeader() As Range
    Dim sh As Worksheet
    Dim anchor As Range
    Dim firstCell As Range
    Dim lastCell As Range

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(CODESHEET)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    On Error Resume Next
        Set anchor = sh.Cells.Find(What:=MODULES_NAME, LookIn:=xlValues, _
                                   LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If anchor Is Nothing Then Exit Function

    Set firstCell = anchor.Offset(1, 0)
    If LenB(Trim$(CStr(firstCell.value))) = 0 Then Exit Function

    Set lastCell = firstCell
    Do While LenB(Trim$(CStr(lastCell.Offset(1, 0).value))) > 0
        Set lastCell = lastCell.Offset(1, 0)
    Loop

    Set FindByHeader = sh.Range(firstCell, lastCell)
End Function

'@sub-title Run one module: ModuleInitialize, each test wrapped in Init/Cleanup, ModuleCleanup.
Private Sub RunModule(ByVal moduleName As String)
    Dim tests As Collection
    Dim idx As Long

    ' Discover FIRST. A module with no '@TestMethod is not a test module -- e.g. a
    ' shared helper (TestHelpersLite) imported only so the real test modules
    ' compile. Skip it entirely: running a lifecycle proc it does not define would
    ' call Application.Run on a missing proc, which -- nested inside the Apple
    ' Events `run VB macro` call -- raises a NON-trappable -50 (RunProc's On Error
    ' Resume Next cannot catch it) that aborts the whole run.
    Set tests = DiscoverTestMethods(moduleName)
    If tests Is Nothing Then Exit Sub
    If tests.Count = 0 Then Exit Sub

    mTestsFound = mTestsFound + tests.Count

    RunProc moduleName, "ModuleInitialize"
    For idx = 1 To tests.Count
        RunProc moduleName, "TestInitialize"
        RunProc moduleName, CStr(tests.Item(idx))
        RunProc moduleName, "TestCleanup"
        DoEvents
    Next idx
    RunProc moduleName, "ModuleCleanup"
End Sub

'@sub-title Application.Run one procedure, swallowing (and logging) any error so the suite continues.
Private Sub RunProc(ByVal moduleName As String, ByVal procName As String)
    If LenB(moduleName) = 0 Or LenB(procName) = 0 Then Exit Sub
    On Error Resume Next
        Application.Run moduleName & "." & procName
        If Err.Number <> 0 Then
            Debug.Print "RunProc " & moduleName & "." & procName & " -> " & _
                        Err.Number & " " & Err.Description
        End If
    On Error GoTo 0
End Sub

'@fun-title Collect the procedure names annotated with '@TestMethod in a module.
Private Function DiscoverTestMethods(ByVal moduleName As String) As Collection
    Dim result As Collection
    Dim component As Object
    Dim codeMod As Object
    Dim lineIndex As Long
    Dim lineText As String
    Dim procName As String

    Set result = New Collection

    ' Probe VBProject access once so the CSV can report whether the "Trust
    ' access to the VBA project object model" setting is blocking discovery.
    On Error Resume Next
        Dim probeCount As Long
        probeCount = ThisWorkbook.VBProject.VBComponents.Count
        If Err.Number <> 0 Then
            mVbeAccess = "ERR" & Err.Number
            Err.Clear
        Else
            mVbeAccess = "ok(" & probeCount & ")"
        End If
    On Error GoTo 0

    On Error Resume Next
        Set component = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not component Is Nothing Then Set codeMod = component.CodeModule
    On Error GoTo 0
    If codeMod Is Nothing Then
        Set DiscoverTestMethods = result
        Exit Function
    End If

    For lineIndex = 1 To codeMod.CountOfLines
        lineText = Trim$(codeMod.Lines(lineIndex, 1))
        If Left$(lineText, 12) = "'@TestMethod" Then
            procName = FindProcedureName(codeMod, lineIndex + 1)
            If LenB(procName) > 0 Then result.Add procName
        End If
    Next lineIndex

    Set DiscoverTestMethods = result
End Function

'@fun-title Find the Sub declared on/after startLine and return its name.
Private Function FindProcedureName(ByVal codeMod As Object, ByVal startLine As Long) As String
    Dim idx As Long
    Dim lineText As String

    For idx = startLine To codeMod.CountOfLines
        lineText = Trim$(codeMod.Lines(idx, 1))
        If LenB(lineText) > 0 Then
            If InStr(1, lineText, "Sub ", vbTextCompare) > 0 Then
                FindProcedureName = ParseProcedureName(lineText)
                Exit Function
            End If
        End If
    Next idx
End Function

'@fun-title Parse the procedure name out of a Sub signature line.
Private Function ParseProcedureName(ByVal signature As String) As String
    Dim tokens() As String
    Dim idx As Long
    Dim candidate As String

    tokens = Split(signature, " ")
    For idx = LBound(tokens) To UBound(tokens)
        candidate = Trim$(tokens(idx))
        If StrComp(candidate, "Sub", vbTextCompare) = 0 Then
            If idx + 1 <= UBound(tokens) Then
                candidate = Trim$(tokens(idx + 1))
                If InStr(candidate, "(") > 0 Then candidate = Left$(candidate, InStr(candidate, "(") - 1)
                ParseProcedureName = candidate
            End If
            Exit Function
        End If
    Next idx
End Function

'@sub-title Serialize testsOutputs to CSV: module,title,type,label,message plus a summary line.
Private Sub SerializeTestOutputs(ByVal path As String)
    Dim sh As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim fileNum As Integer
    Dim title As String
    Dim rawType As String
    Dim testType As String
    Dim label As String
    Dim message As String
    Dim total As Long
    Dim failures As Long

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(TEST_OUTPUT_SHEET)
    On Error GoTo 0

    fileNum = FreeFile
    Open path For Output As #fileNum
    Print #fileNum, "module,title,type,label,message"

    If Not sh Is Nothing Then
        lastRow = LastUsedRow(sh)
        For r = FIRST_DATA_ROW To lastRow
            rawType = CStr(sh.Cells(r, COL_TYPE).value)
            testType = NormalizeType(rawType)
            If LenB(testType) > 0 Then
                title = CleanField(CStr(sh.Cells(r, COL_TITLE).value))
                label = CleanField(CStr(sh.Cells(r, COL_LABEL).value))
                message = CleanField(CStr(sh.Cells(r, COL_MESSAGE).value))
                If LenB(message) = 0 Then message = label

                Print #fileNum, CsvField(title) & "," & CsvField(title) & "," & _
                                CsvField(testType) & "," & CsvField(label) & "," & _
                                CsvField(message)

                total = total + 1
                If testType = "error" Then failures = failures + 1
            End If
        Next r
    End If

    ' Summary line: empty type keeps it out of the success/error tallies in run-tests.R.
    Print #fileNum, CsvField("__summary__") & ",," & CsvField(vbNullString) & "," & _
                    CsvField("total=" & total & " failures=" & failures) & "," & _
                    CsvField("modulesRun=" & mModulesRun & " testsFound=" & mTestsFound & _
                             " vbe=" & mVbeAccess & " name=" & mNameStatus)

    Close #fileNum
End Sub

'@fun-title Last row holding data across cols B:E (rows are sparse, so scan the used range).
Private Function LastUsedRow(ByVal sh As Worksheet) As Long
    Dim used As Range
    Dim lr As Long

    On Error Resume Next
        Set used = sh.UsedRange
        If Not used Is Nothing Then lr = used.Row + used.Rows.Count - 1
    On Error GoTo 0
    If lr < FIRST_DATA_ROW Then lr = FIRST_DATA_ROW
    LastUsedRow = lr
End Function

'@fun-title Map the CheckingOutput type cell ("<symbol> Success|Error|...") to a bare token.
Private Function NormalizeType(ByVal cellText As String) As String
    Dim s As String
    s = Trim$(cellText)
    If LenB(s) = 0 Then Exit Function

    ' Data rows are prefixed with the status glyph CheckingOutput writes.
    Select Case AscW(Left$(s, 1))
        Case 10004: NormalizeType = "success"   ' checkmark
        Case 10060: NormalizeType = "error"     ' cross mark
        Case 9888:  NormalizeType = "warning"
        Case 8505:  NormalizeType = "info"
        Case 9998:  NormalizeType = "note"
    End Select
End Function

'@fun-title Strip embedded newlines so a value never breaks the CSV row structure.
Private Function CleanField(ByVal s As String) As String
    Dim result As String
    result = Replace(s, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    CleanField = result
End Function

'@fun-title Quote a CSV field and escape embedded quotes.
Private Function CsvField(ByVal s As String) As String
    CsvField = """" & Replace(s, """", """""") & """"
End Function

'@fun-title Absolute path for the results CSV, next to the running workbook.
Private Function ResultsPath() As String
    ResultsPath = ThisWorkbook.Path & Application.PathSeparator & RESULTS_FILENAME
End Function
