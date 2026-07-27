Attribute VB_Name = "TestHelpersLite"
Attribute VB_Description = "Minimal test helpers for the OSFiles suite (CustomTest dependency only)"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@Folder("Tests")
'@ModuleDescription("Minimal test helpers for the OSFiles suite (CustomTest dependency only)")

' =============================================================================
' A deliberately small subset of TestHelpers, carrying ONLY the routines that
' TestOSFiles calls: BusyApp / RestoreApp / EnsureWorksheet (+ ClearWorksheet) /
' CustomTestSetTitles / FailUnexpectedError. Its single project dependency is
' CustomTest, which is always present in the baseline.
'
' The full TestHelpers additionally binds TranslationObject and BetterArray
' (BuildTranslationObject, BetterArrayFromList, ...). When a minimal run imports
' only OSFiles + its test, those extra classes may be absent, so the full module
' would not compile and the whole project would fail -- surfacing as the opaque
' AppleScript -50 at OBTRunAllTests. This lite variant removes that coupling so
' the OSFiles suite compiles from just its own dependencies. The full TestHelpers
' stays on disk for suites that need the richer helpers.
'
' The public names mirror TestHelpers exactly, so only one of the two may be
' imported at a time (importing both would raise "Ambiguous name detected").
' The registry registers this lite module for the minimal OSFiles run.
' =============================================================================

'@section Application State
'===============================================================================

'@label BusyApp
'@sub-title Suspend heavy Excel UI features while tests manipulate workbooks.
'@details Suspend heavy Excel UI features while tests manipulate workbooks.
Public Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'@label RestoreApp
'@sub-title Restore the Excel UI to its default behaviour after BusyApp.
'@details Restore the Excel UI to its default behaviour after BusyApp.
Public Sub RestoreApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

'@section Worksheets
'===============================================================================

'@label EnsureWorksheet
'@fun-title Ensure a worksheet exists and is cleared.
'@details Ensure a worksheet exists and is cleared.
'@param sheetName String. Name of the worksheet to create/reset.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Worksheet ensured for use.
Public Function EnsureWorksheet(ByVal sheetName As String, _
                                Optional ByVal targetBook As workbook, _
                                Optional ByVal clearSheet As Boolean = True, _
                                Optional ByVal visibility As Long = xlSheetVisible) As Worksheet

    Dim wb As workbook
    Dim sh As Worksheet

    If (targetBook Is Nothing) Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    If sh Is Nothing Then
        BusyApp
        Set sh = wb.Worksheets.Add
        sh.Name = sheetName
    End If

    sh.Visible = visibility
    If clearSheet Then
        ClearWorksheet sh
    End If

    Set EnsureWorksheet = sh
End Function

'@label ClearWorksheet
'@sub-title Remove data, tables, shapes and names from a worksheet.
'@details Remove data, tables, shapes and names from a worksheet.
'@param sh Worksheet to clear.
Public Sub ClearWorksheet(ByVal sh As Worksheet)

    Dim nm As Name

    If sh Is Nothing Then Exit Sub

    BusyApp

    On Error Resume Next
        Do While sh.ListObjects.Count > 0
            sh.ListObjects(1).Delete
        Loop

        Do While sh.Shapes.Count > 0
            sh.Shapes(1).Delete
        Loop

        For Each nm In sh.Names
            nm.Delete
        Next nm

        For Each nm In sh.Parent.Names
            If InStr(1, nm.RefersTo, "'" & sh.Name & "'!", vbTextCompare) > 0 Then nm.Delete
        Next nm

        sh.Cells.Clear
    On Error GoTo 0
End Sub

'@section Assertions
'===============================================================================

'@label FailUnexpectedError
'@sub-title Fail the current test when unexpected errors surface.
'@details Fail the current test when unexpected errors surface.
'@param assertObj Rubberduck Assert object.
'@param routineName String. Name of the failing routine.
Public Sub FailUnexpectedError(assertObj As Object, ByVal routineName As String)
    On Error Resume Next
    assertObj.Fail "Unexpected error in " & routineName & ": " & Err.Number & " - " & Err.description
    On Error GoTo 0
End Sub

'@label CustomTestSetTitles
'@sub-title Configure the pending test title and subtitle for a CustomTest harness.
'@details Safely sets the next test name and subtitle when the harness reference is valid.
'@param harness CustomTest harness instance.
'@param testName String title to assign.
'@param testSubtitle Optional String subtitle to assign.
Public Sub CustomTestSetTitles(ByVal harness As CustomTest, _
                              ByVal testName As String, _
                              Optional ByVal testSubtitle As String = vbNullString)
    If harness Is Nothing Then Exit Sub
    harness.SetTestName testName
    harness.SetTestSubtitle testSubtitle
End Sub
