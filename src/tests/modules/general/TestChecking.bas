Attribute VB_Name = "TestChecking"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the Checking class")

'@description
'Validates the Checking class, the project-wide log container that stores
'keyed entries with a label, a severity type (Error, Warning, Note, Info,
'Success), and a colour. Tests cover factory creation, name and heading
'properties, entry addition, length tracking, key existence and retrieval,
'value lookup by attribute (label, type, colour), deep-clone independence,
'and the Append method that merges entries from one Checking into another.
'A fresh Checking instance with four entries (one per severity) is created
'in TestInitialize so every test starts from the same baseline.
'Uses the Rubberduck test runner (Rubberduck.AssertClass).
'@depends Checking, IChecking, TestHelpers

Private Assert As Object
Private Fakes As Object
Private CheckingTitle As String
Private CheckingSubTitle As String
Private checkingUnderTest As IChecking

'@section Helpers
'===============================================================================

'@sub-title Populate a checking instance with four entries covering all severity levels.
Private Sub PopulateDefaultEntries(ByVal checkingInstance As IChecking)
    checkingInstance.Add "key-1", "Key 1, error", checkingError
    checkingInstance.Add "key-2", "Key 2, warning", checkingWarning
    checkingInstance.Add "key-3", "Key 3, Note", checkingNote
    checkingInstance.Add "key-4", "Key 4, Info", checkingInfo
End Sub

'@sub-title Strip Unicode severity icons from a type label for comparison.
'@details
'The Checking class prepends emoji icons to type labels. This helper
'removes known icon code points so assertions can compare the plain-text
'severity name (Error, Warning, Note, Info, Success).
Private Function NormaliseTypeLabel(ByVal typeLabel As String) As String
    Dim cleaned As String

    cleaned = typeLabel
    cleaned = Replace(cleaned, ChrW(10060), vbNullString)
    cleaned = Replace(cleaned, ChrW(9888), vbNullString)
    cleaned = Replace(cleaned, ChrW(8505), vbNullString)
    cleaned = Replace(cleaned, ChrW(9998), vbNullString)
    cleaned = Replace(cleaned, ChrW(10004), vbNullString)

    NormaliseTypeLabel = Trim$(cleaned)
End Function

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
    CheckingTitle = "Test Checking title"
    CheckingSubTitle = "Test Checking subtitle"
    Set checkingUnderTest = checking.Create(CheckingTitle, CheckingSubTitle)
    PopulateDefaultEntries checkingUnderTest
End Sub

'@TestCleanUp
Private Sub TestCleanUp()
    Set checkingUnderTest = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify factory creates a valid Checking with title only or title and subtitle.
'@details
'Calls Checking.Create with one argument and then with two. Both calls
'must return a non-Nothing IChecking reference.
'@TestMethod("Checkings")
Public Sub TestCreateCheck()
    On Error GoTo Fail

    Dim checkingInstance As IChecking

    '@
    Set checkingInstance = checking.Create("A title")
    Assert.IsTrue (Not checkingInstance Is Nothing), "Unable to create a checking object with title only"

    Set checkingInstance = checking.Create("A title", "A subtitle")
    Assert.IsTrue (Not checkingInstance Is Nothing), "Unable to create a checking object with title and subtitle"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCreateCheck"
End Sub

'@sub-title Verify the Name property returns the title passed to Create.
'@TestMethod("Checkings")
Public Sub TestNameCheck()
    On Error GoTo Fail

    Assert.AreEqual CheckingTitle, checkingUnderTest.Name, "Name of the checking object is not correctly set"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestNameCheck"
End Sub

'@sub-title Verify Heading returns title by default and subtitle when requested.
'@details
'Heading without arguments should match Name. Heading(True) must return
'the subtitle string provided at creation time.
'@TestMethod("Checkings")
Public Sub TestHeadingsCheck()
    On Error GoTo Fail

    Assert.AreEqual checkingUnderTest.Name, checkingUnderTest.Heading, "Checks name and heading returning different results"
    Assert.AreEqual CheckingSubTitle, checkingUnderTest.Heading(True), "Checks subtitle not correctly returned"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestHeadingsCheck"
End Sub

'@sub-title Verify Add stores entries and Length reflects the count.
'@details
'Creates a new Checking, populates it with four entries via the helper,
'then asserts that Length equals 4.
'@TestMethod("Checkings")
Public Sub TestAddValuesCheck()
    On Error GoTo Fail

    Dim checkingInstance As IChecking

    Set checkingInstance = checking.Create(CheckingTitle, CheckingSubTitle)
    PopulateDefaultEntries checkingInstance
    Assert.IsTrue (checkingInstance.Length = 4), "Expected four entries after adding values"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddValuesCheck"
End Sub

'@sub-title Verify Length returns the number of stored entries.
'@TestMethod("Checkings")
Public Sub TestLength()
    On Error GoTo Fail

    Assert.IsTrue (checkingUnderTest.Length = 4), "Wrong checking length"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestLength"
End Sub

'@sub-title Verify key enumeration, existence checks, and duplicate rejection.
'@details
'Asserts ListOfKeys returns a 1-based BetterArray of length 4, that
'KeyExists finds known keys and rejects absent ones, and that adding a
'duplicate key raises ElementShouldNotExists.
'@TestMethod("Checkings")
Public Sub TestKeysCheck()
    On Error GoTo Fail

    Assert.IsTrue (checkingUnderTest.ListOfKeys.lowerBound = 1), "BetterArray of list of keys lowerbound is not 1"
    Assert.IsTrue (checkingUnderTest.ListOfKeys.Length = 4), "Checking list of keys length is not correct"
    Assert.AreEqual "key-1", checkingUnderTest.ListOfKeys.items(1), "Unable to retrieve the value of the first key item"
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

'@sub-title Verify ValueOf returns correct label, type, and colour for each severity.
'@details
'Retrieves the label, normalised type text, and colour for each of the
'four pre-populated entries. Asserts Error, Warning, Note, and Info are
'stored and retrievable with the correct severity colour mapping.
'@TestMethod("Checkings")
Public Sub TestRetrieveValuesCheck()
    On Error GoTo Fail

    Assert.AreEqual "Key 1, error", checkingUnderTest.ValueOf("key-1", checkingLabel), "Unable to retrieve correct label value"
    Assert.AreEqual "Error", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-1", checkingType)), "Unable to retrieve correct type value for a checking Error"
    Assert.AreEqual "Warning", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-2", checkingType)), "Unable to retrieve correct type value for a checking Warning"
    Assert.AreEqual "Note", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-3", checkingType)), "Unable to retrieve correct type value for a checking Note"
    Assert.AreEqual "Info", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-4", checkingType)), "Unable to retrieve correct type value for a checking Info"
    Assert.AreEqual "purple", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-3", checkingColor)), "Unable to retrieve correct color value for a checking note"
    Assert.AreEqual "grey", NormaliseTypeLabel(checkingUnderTest.ValueOf("key-4", checkingColor)), "Unable to retrieve correct color value for a checking info"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRetrieveValuesCheck"
End Sub

'@sub-title Verify Clone produces a deep, independent copy.
'@details
'Clones an original Checking and confirms length, title, and subtitle are
'copied. Then mutates the clone (add key-clone) and the original (add
'key-original) independently, asserting that neither sees the other's new
'entry. This proves the internal BetterArray storage is not shared.
'@TestMethod("Checkings")
Public Sub TestCloneProducesIndependentCopy()
    On Error GoTo Fail

    Dim original As IChecking
    Dim cloned As IChecking

    Set original = checking.Create(CheckingTitle, CheckingSubTitle)
    PopulateDefaultEntries original

    Set cloned = original.Clone

    Assert.IsTrue (Not cloned Is Nothing), "Clone should return a new checking instance"
    Assert.AreEqual original.Length, cloned.Length, "Clone should carry over existing entries"
    Assert.AreEqual original.Name, cloned.Name, "Clone should copy the checking title"
    Assert.AreEqual original.Heading(True), cloned.Heading(True), "Clone should copy the subtitle"

    cloned.Add "key-clone", "Clone-only entry", checkingError

    Assert.IsTrue (cloned.Length = 5), "Clone should record newly added entries"
    Assert.IsTrue (original.Length = 4), "Original should remain unchanged when clone mutates"
    Assert.IsFalse original.KeyExists("key-clone"), "Original should not see clone-only entries"
    Assert.IsTrue cloned.KeyExists("key-clone"), "Clone should contain its new entry"

    original.Add "key-original", "Original-only entry", checkingWarning

    Assert.IsTrue (original.Length = 5), "Original should record its own new entries"
    Assert.IsTrue (cloned.Length = 5), "Clone should remain unchanged when original mutates"
    Assert.IsFalse cloned.KeyExists("key-original"), "Clone should not see original-only entries"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCloneProducesIndependentCopy"
End Sub

'@sub-title Verify Append merges entries from source into destination.
'@details
'Creates a destination with four entries and a source with two new keys.
'After Append, destination has six entries and source remains at two. Keys
'and labels from the source are verified in the destination with correct
'severity types preserved.
'@TestMethod("Checkings")
Public Sub TestAppendCopiesEntriesFromSource()
    On Error GoTo Fail

    Dim destination As IChecking
    Dim source As IChecking

    Set destination = checking.Create(CheckingTitle, CheckingSubTitle)
    PopulateDefaultEntries destination

    Set source = checking.Create("Source title", "Source subtitle")
    source.Add "src-key-1", "Source key note", checkingNote
    source.Add "src-key-2", "Source key error", checkingError

    destination.Append source

    Assert.IsTrue (6 = destination.Length), "Appending entries should increase destination length"
    Assert.IsTrue (2 = source.Length), "Source length should remain unchanged after append"
    Assert.IsTrue destination.KeyExists("src-key-1"), "Destination should contain appended key src-key-1"
    Assert.IsTrue destination.KeyExists("src-key-2"), "Destination should contain appended key src-key-2"
    Assert.AreEqual "Source key note", destination.ValueOf("src-key-1", checkingLabel), "Destination should copy the source label"
    Assert.AreEqual "Note", NormaliseTypeLabel(destination.ValueOf("src-key-1", checkingType)), "Destination should preserve the source scope (Note)"
    Assert.AreEqual "Error", NormaliseTypeLabel(destination.ValueOf("src-key-2", checkingType)), "Destination should preserve the source scope (Error)"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAppendCopiesEntriesFromSource"
End Sub

'@sub-title Verify Append raises when the source contains a duplicate key.
'@details
'Populates a destination with four entries including key-1, then creates a
'source containing the same key-1. Calling Append should raise
'ElementShouldNotExists. The handler confirms the error number and clears.
'@TestMethod("Checkings")
Public Sub TestAppendRaisesErrorWhenDuplicateKeyExists()
    On Error GoTo HandleDuplicate

    Dim destination As IChecking
    Dim source As IChecking

    Set destination = checking.Create(CheckingTitle, CheckingSubTitle)
    PopulateDefaultEntries destination

    Set source = checking.Create("Source title")
    source.Add "key-1", "Duplicate entry", checkingWarning

    destination.Append source
    Assert.Fail "Expected Append to raise when encountering duplicate key"

    Exit Sub

HandleDuplicate:
    If Err.Number = ProjectError.ElementShouldNotExists Then
        Err.Clear
        Exit Sub
    End If
    FailUnexpectedError Assert, "TestAppendRaisesErrorWhenDuplicateKeyExists"
End Sub
