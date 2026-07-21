Attribute VB_Name = "TestCalculator"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Tests for the Calculator class")

'@description
'Exercises the Calculator probe end to end so the AppleScript test loop runs
'across more than one class. Covers the factory default, a seeded start, add,
'subtract, multiply, and both branches of the guarded SafeDivide.
'@depends Calculator, CustomTest

Private Assert As CustomTest
Private calc As Calculator

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCalculator"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set calc = Calculator.Create()
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
    Set calc = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Calculator")
Public Sub TestDefaultTotalIsZero()
    Assert.SetTestName "TestDefaultTotalIsZero"
    On Error GoTo Fail
    Assert.AreEqual CDbl(0), calc.Total, "A new calculator starts at zero"
    Exit Sub
Fail:
    Assert.LogFailure "TestDefaultTotalIsZero raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestConfiguredStart()
    Assert.SetTestName "TestConfiguredStart"
    On Error GoTo Fail
    Dim seeded As Calculator
    Set seeded = Calculator.Create(10)
    Assert.AreEqual CDbl(10), seeded.Total, "Create seeds the running total"
    Exit Sub
Fail:
    Assert.LogFailure "TestConfiguredStart raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestAddAccumulates()
    Assert.SetTestName "TestAddAccumulates"
    On Error GoTo Fail
    calc.Add 3
    calc.Add 4
    Assert.AreEqual CDbl(7), calc.Total, "Add accumulates into the total"
    Exit Sub
Fail:
    Assert.LogFailure "TestAddAccumulates raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestSubtract()
    Assert.SetTestName "TestSubtract"
    On Error GoTo Fail
    calc.Add 10
    calc.Subtract 4
    Assert.AreEqual CDbl(6), calc.Total, "Subtract lowers the total"
    Exit Sub
Fail:
    Assert.LogFailure "TestSubtract raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestMultiply()
    Assert.SetTestName "TestMultiply"
    On Error GoTo Fail
    calc.Add 5
    calc.Multiply 3
    Assert.AreEqual CDbl(15), calc.Total, "Multiply scales the total"
    Exit Sub
Fail:
    Assert.LogFailure "TestMultiply raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestSafeDivideByNonZero()
    Assert.SetTestName "TestSafeDivideByNonZero"
    On Error GoTo Fail
    calc.Add 20
    Dim applied As Boolean
    applied = calc.SafeDivide(4)
    Assert.IsTrue applied, "SafeDivide reports success for a non-zero divisor"
    Assert.AreEqual CDbl(5), calc.Total, "SafeDivide divides the total"
    Exit Sub
Fail:
    Assert.LogFailure "TestSafeDivideByNonZero raised: " & Err.Description
End Sub

'@TestMethod("Calculator")
Public Sub TestSafeDivideByZeroGuarded()
    Assert.SetTestName "TestSafeDivideByZeroGuarded"
    On Error GoTo Fail
    calc.Add 9
    Dim applied As Boolean
    applied = calc.SafeDivide(0)
    Assert.IsFalse applied, "SafeDivide refuses a zero divisor"
    Assert.AreEqual CDbl(9), calc.Total, "A guarded divide leaves the total unchanged"
    Exit Sub
Fail:
    Assert.LogFailure "TestSafeDivideByZeroGuarded raised: " & Err.Description
End Sub
