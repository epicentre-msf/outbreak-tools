Attribute VB_Name = "TestChecking"

Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private checkingTitle As String
Private checkingSubtitle As String
Private checkingUnderTest As IChecking

'@section Helpers
'===============================================================================

Private Sub PopulateDefaultEntries(ByVal checkingInstance As IChecking)
    checkingInstance.Add "key-1", "Key 1, error", checkingError
    checkingInstance.Add "key-2", "Key 2, warning", checkingWarning
    checkingInstance.Add "key-3", "Key 3, Note", checkingNote
    checkingInstance.Add "key-4", "Key 4, Info", checkingInfo
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    checkingTitle = "Test Checking title"
    checkingSubtitle = "Test Checking subtitle"
    Set checkingUnderTest = Checking.Create(checkingTitle, checkingSubtitle)
    PopulateDefaultEntries checkingUnderTest
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Checkings")
Public Sub TestCreateCheck()
    On Error GoTo Fail

    Dim checkingInstance As IChecking

    '@
    Set checkingInstance = Checking.Create("A title")
    Assert.IsTrue (Not checkingInstance Is Nothing), "Unable to create a checking object with title only"

    Set checkingInstance = Checking.Create("A title", "A subtitle")
    Assert.IsTrue (Not checkingInstance Is Nothing), "Unable to create a checking object with title and subtitle"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCreateCheck"
End Sub

'@TestMethod("Checkings")
Public Sub TestNameCheck()
    On Error GoTo Fail

    Assert.AreEqual checkingTitle, checkingUnderTest.Name, "Name of the checking object is not correctly set"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestNameCheck"
End Sub

'@TestMethod("Checkings")
Public Sub TestHeadingsCheck()
    On Error GoTo Fail

    Assert.AreEqual checkingUnderTest.Name, checkingUnderTest.Heading, "Checks name and heading returning different results"
    Assert.AreEqual checkingSubtitle, checkingUnderTest.Heading(True), "Checks subtitle not correctly returned"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestHeadingsCheck"
End Sub

'@TestMethod("Checkings")
Public Sub TestAddValuesCheck()
    On Error GoTo Fail

    Dim checkingInstance As IChecking

    Set checkingInstance = Checking.Create(checkingTitle, checkingSubtitle)
    PopulateDefaultEntries checkingInstance
    Assert.IsTrue (checkingInstance.Length = 4), "Expected four entries after adding values"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddValuesCheck"
End Sub

'@TestMethod("Checkings")
Public Sub TestLength()
    On Error GoTo Fail

    Assert.IsTrue (checkingUnderTest.Length = 4), "Wrong checking length"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestLength"
End Sub

'@TestMethod("Checkings")
Public Sub TestKeysCheck()
    On Error GoTo Fail

    Assert.IsTrue (checkingUnderTest.ListOfKeys.LowerBound = 1), "BetterArray of list of keys lowerbound is not 1"
    Assert.IsTrue (checkingUnderTest.ListOfKeys.Length = 4), "Checking list of keys length is not correct"
    Assert.AreEqual "key-1", checkingUnderTest.ListOfKeys.Items(1), "Unable to retrieve the value of the first key item"
    Assert.IsTrue checkingUnderTest.KeyExists("key-3"), "key-3 not found in the list of keys"
    Assert.IsFalse checkingUnderTest.KeyExists("key-5"), "key-5 is not added, but mentioned as present"

    checkingUnderTest.Add "key-4", "Key 4, Info", checkingInfo
    Assert.Fail "Expected error when adding existing key not raised"

    Exit Sub

Fail:
    If Err.Number = ProjectError.ElementShouldNotExists Then
        Err.Clear
        Exit Sub
    End If
    FailUnexpectedError Assert, "TestKeysCheck"
End Sub

'@TestMethod("Checkings")
Public Sub TestRetrieveValuesCheck()
    On Error GoTo Fail

    Assert.AreEqual "Key 1, error", checkingUnderTest.ValueOf("key-1", checkingLabel), "Unable to retrieve correct label value"
    Assert.AreEqual "Error", checkingUnderTest.ValueOf("key-1", checkingType), "Unable to retrieve correct type value for a checking Error"
    Assert.AreEqual "Warning", checkingUnderTest.ValueOf("key-2", checkingType), "Unable to retrieve correct type value for a checking Warning"
    Assert.AreEqual "Note", checkingUnderTest.ValueOf("key-3", checkingType), "Unable to retrieve correct type value for a checking Note"
    Assert.AreEqual "Info", checkingUnderTest.ValueOf("key-4", checkingType), "Unable to retrieve correct type value for a checking Info"
    Assert.AreEqual "purple", checkingUnderTest.ValueOf("key-3", checkingColor), "Unable to retrieve correct color value for a checking note"
    Assert.AreEqual "grey", checkingUnderTest.ValueOf("key-4", checkingColor), "Unable to retrieve correct color value for a checking info"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRetrieveValuesCheck"
End Sub
