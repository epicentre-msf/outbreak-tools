Attribute VB_Name = "TestGreeter"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Tests for the Greeter class")

'@description
'Exercises the Greeter probe end to end so the AppleScript test loop has
'something to run. Covers the factory default, a configured name, renaming,
'the plain and loud greetings, and the IsAnonymous guard for blank names.
'@depends Greeter, CustomTest

Private Assert As CustomTest
Private greet As Greeter

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGreeter"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set greet = Greeter.Create()
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
    Set greet = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Greeter")
Public Sub TestDefaultNameIsWorld()
    Assert.SetTestName "TestDefaultNameIsWorld"
    On Error GoTo Fail
    Assert.AreEqual "world", greet.name, "A new greeter defaults to world"
    Exit Sub
Fail:
    Assert.LogFailure "TestDefaultNameIsWorld raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestConfiguredName()
    Assert.SetTestName "TestConfiguredName"
    On Error GoTo Fail
    Dim seeded As Greeter
    Set seeded = Greeter.Create("Ada")
    Assert.AreEqual "Ada", seeded.name, "Create seeds the name"
    Exit Sub
Fail:
    Assert.LogFailure "TestConfiguredName raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestDefaultGreeting()
    Assert.SetTestName "TestDefaultGreeting"
    On Error GoTo Fail
    Assert.AreEqual "Hello, world!", greet.Greeting(), "Greeting wraps the name"
    Exit Sub
Fail:
    Assert.LogFailure "TestDefaultGreeting raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestRenameChangesGreeting()
    Assert.SetTestName "TestRenameChangesGreeting"
    On Error GoTo Fail
    greet.Rename "Ada"
    Assert.AreEqual "Hello, Ada!", greet.Greeting(), "Rename feeds the greeting"
    Exit Sub
Fail:
    Assert.LogFailure "TestRenameChangesGreeting raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestLoudGreeting()
    Assert.SetTestName "TestLoudGreeting"
    On Error GoTo Fail
    Assert.AreEqual "HELLO, WORLD!", greet.LoudGreeting(), "LoudGreeting shouts the greeting"
    Exit Sub
Fail:
    Assert.LogFailure "TestLoudGreeting raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestBlankNameIsAnonymous()
    Assert.SetTestName "TestBlankNameIsAnonymous"
    On Error GoTo Fail
    greet.Rename "   "
    Assert.IsTrue greet.IsAnonymous(), "Whitespace-only name is anonymous"
    Exit Sub
Fail:
    Assert.LogFailure "TestBlankNameIsAnonymous raised: " & Err.Description
End Sub

'@TestMethod("Greeter")
Public Sub TestNamedIsNotAnonymous()
    Assert.SetTestName "TestNamedIsNotAnonymous"
    On Error GoTo Fail
    Assert.IsFalse greet.IsAnonymous(), "A named greeter is not anonymous"
    Exit Sub
Fail:
    Assert.LogFailure "TestNamedIsNotAnonymous raised: " & Err.Description
End Sub
