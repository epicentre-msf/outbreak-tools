Attribute VB_Name = "TestHelpersLite"
Attribute VB_Description = "Minimal test helpers for the OSFiles suite (CustomTest dependency only)"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@Folder("Tests")
'@ModuleDescription("Minimal test helpers for the OSFiles suite (CustomTest dependency only)")

' =============================================================================
' A small subset of TestHelpers, carrying only the routines the registered
' CustomTest suites call: BusyApp / RestoreApp / EnsureWorksheet / ClearWorksheet
' / DeleteWorksheet(s) / FailUnexpectedError / CustomTestSetTitles /
' CustomTestLogFailure / BetterArrayFromList. Project dependencies: CustomTest
' and BetterArray, both present in the baseline workbook.
'
' Registered in place of the full TestHelpers. NB the full module also compiles
' in this workbook -- its TranslationObject/BetterArray deps ARE present -- so
' this split is optional: grow it as suites are added, or point the registry
' back at the full TestHelpers once several suites need the richer helpers.
'
' The public names mirror TestHelpers exactly, so only one of the two may be
' imported at a time (importing both would raise "Ambiguous name detected").
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

'@label DeleteWorksheet
'@sub-title Delete a worksheet if it exists.
'@param sheetName String. Worksheet to delete.
Public Sub DeleteWorksheet(ByVal sheetName As String)
    On Error Resume Next
        BusyApp
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

'@label DeleteWorksheets
'@sub-title Delete several worksheets in a single call.
'@param sheetNames ParamArray list of worksheet names.
Public Sub DeleteWorksheets(ParamArray sheetNames() As Variant)
    Dim idx As Long

    For idx = LBound(sheetNames) To UBound(sheetNames)
        DeleteWorksheet CStr(sheetNames(idx))
    Next idx
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

'@label CustomTestLogFailure
'@sub-title Log a formatted failure message on a CustomTest harness.
'@details Builds a descriptive failure message containing the routine name and optional error info, then logs it.
'@param harness CustomTest harness instance.
'@param routineName String name of the failing routine.
'@param errNumber Optional Long error number to include.
'@param errDescription Optional String error description to include.
Public Sub CustomTestLogFailure(ByVal harness As CustomTest, _
                                ByVal routineName As String, _
                                Optional ByVal errNumber As Long = 0, _
                                Optional ByVal errDescription As String = vbNullString)
    Dim message As String
    Dim errorExplanation As String

    If harness Is Nothing Then Exit Sub
    message = routineName

    If errNumber <> 0 Or LenB(errDescription) > 0 Then

        Select Case errNumber
        Case 1001: errorExplanation = "Invalid argument"
        Case 1002: errorExplanation = "Object not initialized"
        Case 1004: errorExplanation = "Unexpected state"
        Case 1005: errorExplanation = "Element should exists"
        Case 1006: errorExplanation = "Element should not exists"
        Case 1007: errorExplanation = "Element not found"
        Case 1008: errorExplanation = "Something went wrong"
        Case Else: errorExplanation = "Unkown error: (" & errNumber & ")"
        End Select

        message = message & ": " & errorExplanation & " - " & errDescription
    End If

    harness.LogFailure message
End Sub

'@section Data Builders
'===============================================================================

'@label BetterArrayFromList
'@fun-title Create a BetterArray with the supplied items.
'@details Create a BetterArray with the supplied items.
'@param items ParamArray values to push.
'@return BetterArray containing the items.
Public Function BetterArrayFromList(ParamArray items() As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.lowerBound = 0

    For idx = LBound(items) To UBound(items)
        result.Push items(idx)
    Next idx

    Set BetterArrayFromList = result
End Function
