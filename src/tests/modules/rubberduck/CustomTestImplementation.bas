Attribute VB_Name = "CustomTestImplementation"
Attribute VB_Description = "Entry point coordinating Rubberduck-style test modules"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("Entry point coordinating Rubberduck-style test modules")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Const DEFAULT_TEST_PREFIX As String = "test"
Private Const DEFAULT_TEST_LONAME As String  = "ModulesForTesting"
Private Const CODESHEET   As String = "Codes"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@EntryPoint
'@label RunModuleTests
'@sub-title Execute all tests found in a module
'@details Runs lifecycle procedures and invokes every `@TestMethod` within the target module. Harness coordination is the responsibility of the test module.
'@param moduleName String name of the module to execute.
Public Sub clickRibbonTests(ByRef Control As IRibbonControl) 

    Dim loRng As Range
    Dim counter As Long
    Dim moduleName As String
    Dim nm As Name

    'Load to keep the screen cool

    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    ThisWorkbook.Worksheets(TEST_OUTPUT_SHEET).Cells.Clear
    ThisWorkbook.Worksheets(TEST_OUTPUT_SHEET).Activate
    Set nm = ThisWorkbook.Worksheets(CODESHEET).Names(DEFAULT_TEST_LONAME)
    
    On Error Resume Next
        Set loRng = nm.RefersToRange
    On Error GoTo 0

    If loRng Is Nothing Then Exit Sub

    On Error GoTo CleanExit
    For counter = 1 To loRng.Rows.Count
        moduleName = loRng.Cells(counter, 1).Value
        ExecuteModule moduleName
    Next
    
CleanExit:
    If Err.Number <> 0 Then
        Debug.Print "Not 0 exit status: Error " & Err.Number & " - Description: " & Err.Description
    Else
        MsgBox "Done!"
    End If

End Sub


'@label ExecuteModule
'@sub-title Internal helper orchestrating a single module run
'@details Handles lifecycle invocations and test discovery for the requested module.
'@param moduleName String module identifier.
Public Sub ExecuteModule(ByVal modName As String)
    Dim tests As BetterArray
    Dim index As Long
    Dim testInfo As Variant
    Dim moduleName As String

    moduleName = Trim$(modName)
    If LenB(moduleName) = 0 Then Exit Sub

    BusyApp
    SafeRun moduleName, "ModuleInitialize"
    DoEvents

    Set tests = DiscoverTests(moduleName)
    If Not tests Is Nothing Then
        For index = tests.LowerBound To tests.UpperBound
            testInfo = tests.Item(index)
            RunSingleTest moduleName, testInfo
            DoEvents
        Next index
    End If

    DoEvents
    SafeRun moduleName, "ModuleCleanup"
End Sub

'@label RunSingleTest
'@sub-title Execute a single test procedure within a module
'@details Runs lifecycle hooks and executes the discovered procedure that represents the test case.
'@param moduleName String module identifier containing the test.
'@param testInfo Variant array emitted by `DiscoverTests` (procedure name, annotation description).
Private Sub RunSingleTest(ByVal moduleName As String, _
                           testInfo As Variant)
    Dim routineName As String

    If IsArray(testInfo) Then
        If UBound(testInfo) >= 0 Then routineName = CStr(testInfo(1))
    Else
        routineName = CStr(testInfo)
    End If

    '@Ignore AssignmentNotUsed
    routineName = Trim$(routineName)
    If LenB(routineName) = 0 Then Exit Sub

    SafeRun moduleName, "TestInitialize"
    DoEvents
    SafeRun moduleName, routineName
    DoEvents
    SafeRun moduleName, "TestCleanup"
End Sub


'@label NormaliseModuleName
'@fun-title Convert a BetterArray entry to a trimmed module name
'@details Handles variants, Empty, or Null values gracefully.
Private Function NormaliseModuleName( value As Variant) As String
    On Error GoTo CleanValue
    If IsEmpty(value) Then Exit Function
    If IsNull(value) Then Exit Function
    NormaliseModuleName = Trim$(CStr(value))
    Exit Function

CleanValue:
    NormaliseModuleName = vbNullString
End Function

'@label SafeRun
'@sub-title Execute a module procedure while suppressing runtime errors
Private Sub SafeRun(ByVal moduleName As String, _
                    ByVal procedureName As String)
    If LenB(moduleName) = 0 Then Exit Sub
    If LenB(procedureName) = 0 Then Exit Sub

    Application.Run moduleName & "." & procedureName
End Sub

'@label DiscoverTests
'@fun-title Inspect a module for `@TestMethod` annotations
Private Function DiscoverTests(ByVal moduleName As String) As BetterArray
    Dim component As Object
    Dim codeMod As Object
    Dim lineIndex As Long
    Dim lineText As String
    Dim tests As BetterArray
    Dim procedureName As String
    Dim description As String

    On Error Resume Next
        Set component = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    If component Is Nothing Then Exit Function
    If StrComp(TypeName(component), "VBComponent", vbTextCompare) <> 0 Then Exit Function

    On Error Resume Next
        Set codeMod = component.CodeModule
    On Error GoTo 0
    If codeMod Is Nothing Then Exit Function
    If StrComp(TypeName(codeMod), "CodeModule", vbTextCompare) <> 0 Then Exit Function
    Set tests = New BetterArray
    tests.LowerBound = 1

    For lineIndex = 1 To codeMod.CountOfLines
        lineText = Trim$(codeMod.Lines(lineIndex, 1))
        If Left$(lineText, 12) = "'@TestMethod" Then
            description = ExtractDescription(lineText)
            procedureName = FindProcedureName(codeMod, lineIndex + 1)
            If LenB(procedureName) > 0 Then
                tests.Push Array(procedureName, description)
            End If
        End If
    Next lineIndex

    If tests.Length = 0 Then Set tests = Nothing
    Set DiscoverTests = tests
End Function

'@label ExtractDescription
'@fun-title Extract the description from a `@TestMethod` annotation
Private Function ExtractDescription(ByVal annotation As String) As String
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(annotation, "(")
    endPos = InStr(annotation, ")")
    If startPos > 0 And endPos > startPos Then
        ExtractDescription = Mid$(annotation, startPos + 1, endPos - startPos - 1)
    Else
        ExtractDescription = DEFAULT_TEST_PREFIX
    End If
End Function

'@label FindProcedureName
'@fun-title Determine the procedure declared after an annotation line
Private Function FindProcedureName(ByVal codeMod As Object, _
                                   ByVal startLine As Long) As String
    Dim idx As Long
    Dim lineText As String

    If codeMod Is Nothing Then Exit Function
    If StrComp(TypeName(codeMod), "CodeModule", vbTextCompare) <> 0 Then Exit Function

    For idx = startLine To codeMod.CountOfLines
        lineText = Trim$(codeMod.Lines(idx, 1))
        If LenB(lineText) = 0 Then GoTo ContinueLoop
        If InStr(1, lineText, "Sub", vbTextCompare) > 0 Then
            FindProcedureName = ParseProcedureName(lineText)
            Exit Function
        End If
ContinueLoop:
    Next idx
End Function

'@label ParseProcedureName
'@fun-title Parse a procedure name from its signature line
Private Function ParseProcedureName(ByVal signature As String) As String
    Dim tokens() As String
    Dim idx As Long
    Dim candidate As String

    tokens = Split(signature, " ")
    For idx = LBound(tokens) To UBound(tokens)
        candidate = Trim$(tokens(idx))
        If LenB(candidate) = 0 Then GoTo ContinueLoop
        If StrComp(candidate, "Sub", vbTextCompare) = 0 Then
            If idx + 1 <= UBound(tokens) Then
                candidate = Trim$(tokens(idx + 1))
                If InStr(candidate, "(") > 0 Then
                    candidate = Left$(candidate, InStr(candidate, "(") - 1)
                End If
                ParseProcedureName = candidate
            End If
            Exit Function
        End If
ContinueLoop:
    Next idx
End Function
