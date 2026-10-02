Attribute VB_Name = "TestLinelist"
Attribute VB_Description = "Tests for Linelist class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for Linelist class")

Option Explicit

'@description
'Drives Linelist, the facade over the output workbook a generation produces.
'Prepare, SaveLL and PrepareAdmin need a real designer, a setup file and a
'template on disk, so what this suite reaches is everything below them: the
'worksheet accessors, the scoped-name cut that decides what a sheet is called,
'the two managers whose state the generation report reads, the discard that
'keeps a stopped build on disk, and the guard that stops a build over
'unprepared specifications.
'
'THE FIXTURE IS BUILT ONCE PER MODULE
'-------------------------------------------------------------------------------
'Two workbooks: one shaped like a designer, carrying the seven worksheets
'LinelistSpecs.Create validates and a dictionary fixture, and one standing in for
'the linelist. The second reaches the class through
'LinelistSpecs.TestAssignLLWorkbook, which is the field Prepare writes in
'production. Building both per test costs minutes over twenty tests.
'
'Each test creates its own Linelist over the shared specifications, because the
'report entries and the cached managers are per instance.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'whole headless run comes back with no results file. The setup captures the
'error into two module fields instead and a guard at the top of every test
'reports it as that test's own failure.
'
'WHAT THE SCOPED NAME TESTS WATCH
'-------------------------------------------------------------------------------
'A companion worksheet is the base name under a prefix, and Excel accepts 31
'characters. The cut used to be applied when a sheet was created and skipped when
'one was looked up, so a base name of 25 characters or more created one sheet and
'asked for another, and the build died.
'TestALongPrintNameResolvesToTheSheetThatWasCreated is what holds the two sides
'together, and TestALongCRFNameKeepsItsWholeName states that the same base name
'under the shorter prefix is left as it is.
'@depends Linelist, LinelistSpecs, LLdictionary, LLSheets, DropdownLists,
'  CustomPivotTable, CustomTest, DictionaryTestFixture

Private Assert As CustomTest
Private SpecsWkb As Workbook
Private OutWkb As Workbook
Private Specs As LinelistSpecs
Private Dict As LLdictionary
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "Linelist"
Private Const DICTIONARY_SHEET As String = "DictFixture"

'The sheets Dropdown and Pivots resolve. The translator the fixture installs
'answers every tag with the tag itself, so these are the names the class asks for.
Private Const SHEET_CUSTOM_CHOICE As String = "LLSHEET_CustomChoice"
Private Const SHEET_CUSTOM_PIVOT As String = "LLSHEET_CustomPivotTable"
Private Const SHEET_DROPDOWN_LISTS As String = "__dropdown_lists"

'26 characters. With the print_ prefix that is 32, one over what Excel accepts,
'and with crf_ it is 30, which fits. The two together state that the cut happens
'when the whole name is too long rather than when the base name is.
Private Const LONG_BASE_NAME As String = "abcdefghijklmnopqrstuvwxyz"

'11 characters, so no scope of it ever needs cutting.
Private Const SHORT_BASE_NAME As String = "shortEnough"

Private Const SHEET_NAME_LIMIT As Long = 31


'@section Lifecycle
'===============================================================================

'@sub-title Build the two workbooks and the specifications, once.
'@details
'This routine is Public because the harness reaches it through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLinelist"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set SpecsWkb = NewWorkbook()
        Set OutWkb = NewWorkbook()

        DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, SpecsWkb
        Set Dict = LLdictionary.Create(SpecsWkb.Worksheets(DICTIONARY_SHEET), 1, 1)
        Dict.Prepare

        EnsureSpecsSheets SpecsWkb

        'The three worksheets the two cached managers read
        EnsureWorksheet SHEET_DROPDOWN_LISTS, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
        EnsureWorksheet SHEET_CUSTOM_CHOICE, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
        EnsureWorksheet SHEET_CUSTOM_PIVOT, OutWkb, clearSheet:=True, visibility:=xlSheetHidden

        Set Specs = LinelistSpecs.Create(SpecsWkb)
        Specs.TestAssignDictionary Dict
        Specs.TestAssignTransObject BuildTranslationObject(SpecsWkb, "ENG", Array())
        Specs.TestAssignLLWorkbook OutWkb

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print results and drop the two workbooks.
'@details
'This routine is Public because the harness reaches it through Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    On Error Resume Next
        If Not SpecsWkb Is Nothing Then DeleteWorkbook SpecsWkb
        If Not OutWkb Is Nothing Then DeleteWorkbook OutWkb
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set SpecsWkb = Nothing
    Set OutWkb = Nothing

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending at that moment, and the Flush in TestCleanup has
'just reset those to the default, so every result of the module would be filed
'under the default label. Letting the first assertion of each test open the
'checking picks up the titles CustomTestSetTitles set at the top of the test.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function


'@sub-title Drop the output folder entry and the files the discard test wrote.
'@param keepFolder String. The folder the entry named.
'@param keptPath String. The kept file, or empty.
Private Sub DiscardFixtureFolder(ByVal keepFolder As String, ByVal keptPath As String)
    On Error Resume Next
    SpecsWkb.Names("RNG_LLDir").Delete
    SpecsWkb.Worksheets("Main").Range("A1").ClearContents
    If LenB(keptPath) > 0 Then Kill keptPath
    If LenB(keepFolder) > 0 Then
        RmDir keepFolder & Application.PathSeparator & "OBTApp_"
        RmDir keepFolder
    End If
    On Error GoTo 0
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestCreateReturnsALinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsALinelist"
    If Not FixtureReady("TestCreateReturnsALinelist") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsALinelist", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestCreateRejectsNothingSpecifications()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSpecifications"
    On Error GoTo ExpectError

    Dim sut As Linelist
    Set sut = Linelist.Create(Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecifications", , _
                         "Expected error when specs is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when specs is Nothing"
End Sub


'@TestMethod("Linelist")
Public Sub TestDiscardBuildClosesTheOutputWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestDiscardBuildClosesTheOutputWorkbook"
    If Not FixtureReady("TestDiscardBuildClosesTheOutputWorkbook") Then Exit Sub
    On Error GoTo TestFail

    'Arrange: own specifications over the shared designer fixture, with a
    'throwaway workbook standing in for the output. The shared Specs and
    'OutWkb stay untouched, so the other tests keep their fixture.
    Dim tempWkb As Workbook
    Dim tempName As String
    Set tempWkb = NewWorkbook()
    tempName = tempWkb.Name

    Dim ownSpecs As LinelistSpecs
    Set ownSpecs = LinelistSpecs.Create(SpecsWkb)
    ownSpecs.TestAssignLLWorkbook tempWkb

    Dim sut As Linelist
    Set sut = Linelist.Create(ownSpecs)

    'Act. The fixture Main carries no output folder, so there is no
    'repository to keep the file in.
    Dim keptPath As String
    keptPath = sut.DiscardBuild()

    'Closing a workbook hands the screen to whichever workbook Excel
    'activates, and PrintResults raises 1004 when the driver workbook has
    'lost it. Hand it back right after the close.
    ThisWorkbook.Activate

    'Assert: the workbook is closed and both references are dropped
    Assert.AreEqual vbNullString, keptPath, _
                    "With no output folder known DiscardBuild should keep nothing"
    Assert.IsTrue ownSpecs.LLWorkbook Is Nothing, _
                  "DiscardBuild should release the specifications' workbook reference"

    Dim stillOpen As Boolean
    On Error Resume Next
    stillOpen = LenB(Application.Workbooks(tempName).Name) > 0
    On Error GoTo TestFail
    Assert.IsFalse stillOpen, _
                   "DiscardBuild should close the output workbook with its changes thrown away"

    Exit Sub
TestFail:
    'An On Error statement clears Err, so the fault is captured first
    Dim failNumber As Long
    Dim failDesc As String
    failNumber = Err.Number
    failDesc = Err.Description

    On Error Resume Next
    If Not tempWkb Is Nothing Then tempWkb.Close savechanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestDiscardBuildClosesTheOutputWorkbook", failNumber, failDesc
End Sub

'@TestMethod("Linelist")
Public Sub TestDiscardBuildWithNoWorkbookIsQuiet()
    CustomTestSetTitles Assert, TESTMODULE, "TestDiscardBuildWithNoWorkbookIsQuiet"
    If Not FixtureReady("TestDiscardBuildWithNoWorkbookIsQuiet") Then Exit Sub
    On Error GoTo TestFail

    'Own specifications with no output workbook: the state before Prepare,
    'and the state after SaveLL already closed the file
    Dim ownSpecs As LinelistSpecs
    Set ownSpecs = LinelistSpecs.Create(SpecsWkb)

    Dim sut As Linelist
    Set sut = Linelist.Create(ownSpecs)

    'Act: with no workbook the call exits quietly
    Dim keptPath As String
    keptPath = sut.DiscardBuild()

    Assert.AreEqual vbNullString, keptPath, _
                    "With no workbook DiscardBuild should answer an empty path"
    Assert.IsTrue ownSpecs.LLWorkbook Is Nothing, _
                  "DiscardBuild should leave the empty reference as it is"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDiscardBuildWithNoWorkbookIsQuiet", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestDiscardBuildKeepsTheFileWhenTheFolderIsKnown()
    CustomTestSetTitles Assert, TESTMODULE, "TestDiscardBuildKeepsTheFileWhenTheFolderIsKnown"
    If Not FixtureReady("TestDiscardBuildKeepsTheFileWhenTheFolderIsKnown") Then Exit Sub
    On Error GoTo TestFail

    'Arrange: an output folder entry on the fixture Main, the way a filled
    'designer carries one, and a throwaway workbook standing in for the
    'output. The name is dropped again below so the other tests keep a
    'Main with no folder.
    Dim keepFolder As String
    Dim mainSheet As Worksheet
    Dim tempWkb As Workbook
    Dim tempName As String
    Dim expectedPath As String
    Dim keptPath As String

    keepFolder = BuildTempFolder(Nothing, "obt_linelist_discard")
    Set mainSheet = SpecsWkb.Worksheets("Main")
    mainSheet.Range("A1").Value = keepFolder
    SpecsWkb.Names.Add Name:="RNG_LLDir", RefersTo:=mainSheet.Range("A1")

    Set tempWkb = NewWorkbook()
    tempName = tempWkb.Name

    Dim ownSpecs As LinelistSpecs
    Set ownSpecs = LinelistSpecs.Create(SpecsWkb)
    ownSpecs.TestAssignLLWorkbook tempWkb

    Dim sut As Linelist
    Set sut = Linelist.Create(ownSpecs)

    expectedPath = keepFolder & Application.PathSeparator & "OBTApp_" & _
                   Application.PathSeparator & "__temp.xlsb"

    'Act
    keptPath = sut.DiscardBuild()
    ThisWorkbook.Activate

    'Assert: the file is where the report will name it, and the workbook is closed
    Assert.AreEqual expectedPath, keptPath, _
                    "DiscardBuild should keep the workbook as __temp.xlsb in the repository"
    Assert.IsTrue LenB(Dir$(keptPath)) > 0, _
                  "The kept file should be on disk"
    Assert.IsTrue ownSpecs.LLWorkbook Is Nothing, _
                  "DiscardBuild should release the specifications' workbook reference"

    Dim stillOpen As Boolean
    On Error Resume Next
    stillOpen = LenB(Application.Workbooks(tempName).Name) > 0
    On Error GoTo TestFail
    Assert.IsFalse stillOpen, _
                   "DiscardBuild should close the output workbook once it is kept"

    'The kept file survives the drop of the linelist and its repository
    Set sut = Nothing
    Set ownSpecs = Nothing
    Assert.IsTrue LenB(Dir$(keptPath)) > 0, _
                  "The kept file should survive the drop of the repository"

    DiscardFixtureFolder keepFolder, keptPath
    Exit Sub
TestFail:
    'An On Error statement clears Err, so the fault is captured first
    Dim failNumber As Long
    Dim failDesc As String
    failNumber = Err.Number
    failDesc = Err.Description

    On Error Resume Next
    If Not tempWkb Is Nothing Then tempWkb.Close savechanges:=False
    ThisWorkbook.Activate
    DiscardFixtureFolder keepFolder, keptPath
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestDiscardBuildKeepsTheFileWhenTheFolderIsKnown", failNumber, failDesc
End Sub


'@section Collaborator tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestLinelistDataReturnsSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestLinelistDataReturnsSpecs"
    If Not FixtureReady("TestLinelistDataReturnsSpecs") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue sut.LinelistData Is Specs, _
                  "LinelistData should answer the specifications it was built with"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLinelistDataReturnsSpecs", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestDictionaryReturnsDictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestDictionaryReturnsDictionary"
    If Not FixtureReady("TestDictionaryReturnsDictionary") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue sut.Dictionary Is Dict, _
                  "Dictionary should answer the dictionary held by the specifications"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDictionaryReturnsDictionary", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestSheetNamesIsWalkedOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesIsWalkedOnce"
    If Not FixtureReady("TestSheetNamesIsWalkedOnce") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim firstRead As BetterArray
    Dim secondRead As BetterArray

    Set sut = Linelist.Create(Specs)
    Set firstRead = sut.SheetNames
    Set secondRead = sut.SheetNames

    Assert.IsTrue firstRead.Length > 0, _
                  "The dictionary fixture should carry at least one sheet name"
    Assert.IsTrue firstRead Is secondRead, _
                  "The second read should answer the list the first read walked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesIsWalkedOnce", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestSheetInfoManagerIsBuiltOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetInfoManagerIsBuiltOnce"
    If Not FixtureReady("TestSheetInfoManagerIsBuiltOnce") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim firstRead As LLSheets
    Dim secondRead As LLSheets

    Set sut = Linelist.Create(Specs)
    Set firstRead = sut.SheetInfoManager
    Set secondRead = sut.SheetInfoManager

    Assert.IsTrue Not firstRead Is Nothing, _
                  "SheetInfoManager should answer an instance"
    Assert.IsTrue firstRead Is secondRead, _
                  "The second read should answer the instance the first read built"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetInfoManagerIsBuiltOnce", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestSheetInfoManagerReadsTheSpecsDictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetInfoManagerReadsTheSpecsDictionary"
    If Not FixtureReady("TestSheetInfoManagerReadsTheSpecsDictionary") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue sut.SheetInfoManager.Dictionary Is Dict, _
                  "The manager should read the dictionary held by the specifications"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetInfoManagerReadsTheSpecsDictionary", Err.Number, Err.Description
End Sub


'@section Worksheet tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestSheetExistsAnswersTheOutputWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetExistsAnswersTheOutputWorkbook"
    If Not FixtureReady("TestSheetExistsAnswersTheOutputWorkbook") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue sut.SheetExists(SHEET_DROPDOWN_LISTS), _
                  "SheetExists should find a sheet the output workbook carries"
    Assert.IsTrue Not sut.SheetExists("NonExistentSheet__xyz"), _
                  "SheetExists should answer False for a sheet that is absent"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetExistsAnswersTheOutputWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestAddOutputSheetCreatesTheSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddOutputSheetCreatesTheSheet"
    If Not FixtureReady("TestAddOutputSheetCreatesTheSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim sheetName As String

    sheetName = "addedByAddOutput"
    Set sut = Linelist.Create(Specs)
    sut.AddOutputSheet sheetName

    Assert.IsTrue sut.SheetExists(sheetName), _
                  "AddOutputSheet should create the worksheet"
    Assert.IsTrue sut.Wksh(sheetName).Visible = xlSheetVeryHidden, _
                  "AddOutputSheet should hide the sheet it creates by default"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddOutputSheetCreatesTheSheet", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestAddOutputSheetLeavesAnExistingSheetAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddOutputSheetLeavesAnExistingSheetAlone"
    If Not FixtureReady("TestAddOutputSheetLeavesAnExistingSheetAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim sheetName As String
    Dim countBefore As Long

    sheetName = "addedTwiceOnPurpose"
    Set sut = Linelist.Create(Specs)
    sut.AddOutputSheet sheetName
    sut.Wksh(sheetName).Range("A1").Value = "kept"

    countBefore = OutWkb.Worksheets.Count
    sut.AddOutputSheet sheetName

    Assert.IsTrue OutWkb.Worksheets.Count = countBefore, _
                  "A second add of the same name should create no worksheet"
    Assert.IsTrue sut.Wksh(sheetName).Range("A1").Value = "kept", _
                  "A second add should leave what the sheet already holds"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddOutputSheetLeavesAnExistingSheetAlone", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestAddOutputSheetInsertsBeforeTheAnchor()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddOutputSheetInsertsBeforeTheAnchor"
    If Not FixtureReady("TestAddOutputSheetInsertsBeforeTheAnchor") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim anchor As Worksheet
    Dim inserted As Worksheet

    Set sut = Linelist.Create(Specs)
    sut.AddOutputSheet "anchorSheet"
    Set anchor = sut.Wksh("anchorSheet")

    sut.AddOutputSheet "insertedSheet", beforeSheet:=anchor
    Set inserted = sut.Wksh("insertedSheet")

    Assert.IsTrue inserted.Index = anchor.Index - 1, _
                  "The new sheet should sit immediately before the anchor"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddOutputSheetInsertsBeforeTheAnchor", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestPrintWkshFindsThePrefixedSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestPrintWkshFindsThePrefixedSheet"
    If Not FixtureReady("TestPrintWkshFindsThePrefixedSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    sut.AddOutputSheet SHORT_BASE_NAME, xlSheetVeryHidden, sheetScope:=2

    Assert.IsTrue sut.PrintWksh(SHORT_BASE_NAME).Name = "print_" & SHORT_BASE_NAME, _
                  "PrintWksh should find the sheet AddOutputSheet made under the print_ prefix"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPrintWkshFindsThePrefixedSheet", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestCRFWkshFindsThePrefixedSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCRFWkshFindsThePrefixedSheet"
    If Not FixtureReady("TestCRFWkshFindsThePrefixedSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    sut.AddOutputSheet SHORT_BASE_NAME, xlSheetVeryHidden, sheetScope:=3

    Assert.IsTrue sut.CRFWksh(SHORT_BASE_NAME).Name = "crf_" & SHORT_BASE_NAME, _
                  "CRFWksh should find the sheet AddOutputSheet made under the crf_ prefix"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCRFWkshFindsThePrefixedSheet", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestWkshRaisesForAMissingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestWkshRaisesForAMissingSheet"
    If Not FixtureReady("TestWkshRaisesForAMissingSheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As Linelist
    Dim sh As Worksheet

    Set sut = Linelist.Create(Specs)
    Set sh = sut.Wksh("thisSheetIsNotThere")

    CustomTestLogFailure Assert, "TestWkshRaisesForAMissingSheet", , _
                         "Expected an error for a sheet the workbook does not carry"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Wksh should raise for a sheet the output workbook does not carry"
End Sub


'@section Scoped name tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestALongPrintNameResolvesToTheSheetThatWasCreated()
    CustomTestSetTitles Assert, TESTMODULE, "TestALongPrintNameResolvesToTheSheetThatWasCreated"
    If Not FixtureReady("TestALongPrintNameResolvesToTheSheetThatWasCreated") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim expectedName As String

    expectedName = Left$("print_" & LONG_BASE_NAME, SHEET_NAME_LIMIT)

    Set sut = Linelist.Create(Specs)
    sut.AddOutputSheet LONG_BASE_NAME, xlSheetVeryHidden, sheetScope:=2

    Assert.IsTrue Len(expectedName) = SHEET_NAME_LIMIT, _
                  "The fixture name should be one Excel has to shorten"
    Assert.IsTrue sut.PrintWksh(LONG_BASE_NAME).Name = expectedName, _
                  "PrintWksh should find the shortened sheet AddOutputSheet made"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongPrintNameResolvesToTheSheetThatWasCreated", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestALongCRFNameKeepsItsWholeName()
    CustomTestSetTitles Assert, TESTMODULE, "TestALongCRFNameKeepsItsWholeName"
    If Not FixtureReady("TestALongCRFNameKeepsItsWholeName") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    sut.AddOutputSheet LONG_BASE_NAME, xlSheetVeryHidden, sheetScope:=3

    Assert.IsTrue sut.CRFWksh(LONG_BASE_NAME).Name = "crf_" & LONG_BASE_NAME, _
                  "The crf_ prefix keeps this name inside the limit, so nothing is cut"
    Assert.IsTrue Not sut.HasCheckings, _
                  "A name that fits should file no report entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongCRFNameKeepsItsWholeName", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestAShortenedSheetNameIsReported()
    CustomTestSetTitles Assert, TESTMODULE, "TestAShortenedSheetNameIsReported"
    If Not FixtureReady("TestAShortenedSheetNameIsReported") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    sut.AddOutputSheet LONG_BASE_NAME, xlSheetVeryHidden, sheetScope:=2

    Assert.IsTrue sut.HasCheckings, _
                  "A name the cut shortened should be reported"
    Assert.IsTrue sut.CheckingValues.Length = 1, _
                  "One shortened name should file one report entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAShortenedSheetNameIsReported", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestAShortenedSheetNameIsReportedOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestAShortenedSheetNameIsReportedOnce"
    If Not FixtureReady("TestAShortenedSheetNameIsReportedOnce") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim sh As Worksheet

    Set sut = Linelist.Create(Specs)

    sut.AddOutputSheet LONG_BASE_NAME, xlSheetVeryHidden, sheetScope:=2
    sut.AddOutputSheet LONG_BASE_NAME, xlSheetVeryHidden, sheetScope:=2
    Set sh = sut.PrintWksh(LONG_BASE_NAME)

    Assert.IsTrue sut.CheckingValues.Length = 1, _
                  "The same shortened name should be reported once whatever asks for it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAShortenedSheetNameIsReportedOnce", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestCheckingValuesIsNothingUntilSomethingIsFiled()
    CustomTestSetTitles Assert, TESTMODULE, "TestCheckingValuesIsNothingUntilSomethingIsFiled"
    If Not FixtureReady("TestCheckingValuesIsNothingUntilSomethingIsFiled") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut.HasCheckings, _
                  "A fresh instance should have nothing to report"
    Assert.IsTrue sut.CheckingValues Is Nothing, _
                  "CheckingValues should answer Nothing while nothing has been filed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCheckingValuesIsNothingUntilSomethingIsFiled", Err.Number, Err.Description
End Sub


'@section Cached manager tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestDropdownKeepsTheInstanceItBuilt()
    CustomTestSetTitles Assert, TESTMODULE, "TestDropdownKeepsTheInstanceItBuilt"
    If Not FixtureReady("TestDropdownKeepsTheInstanceItBuilt") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim firstRead As DropdownLists
    Dim secondRead As DropdownLists

    Set sut = Linelist.Create(Specs)
    Set firstRead = sut.Dropdown(1)
    Set secondRead = sut.Dropdown(1)

    Assert.IsTrue firstRead Is secondRead, _
                  "Dropdown should answer the same manager, so what it records survives"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDropdownKeepsTheInstanceItBuilt", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestTheTwoDropdownScopesAreTwoManagers()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheTwoDropdownScopesAreTwoManagers"
    If Not FixtureReady("TestTheTwoDropdownScopesAreTwoManagers") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim standard As DropdownLists
    Dim custom As DropdownLists

    Set sut = Linelist.Create(Specs)
    Set standard = sut.Dropdown(1)
    Set custom = sut.Dropdown(2)

    Assert.IsTrue Not standard Is custom, _
                  "Each scope should carry its own manager"
    Assert.IsTrue sut.Dropdown(2) Is custom, _
                  "The custom scope should keep the manager it built"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTwoDropdownScopesAreTwoManagers", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestDropdownRejectsAnUnknownScope()
    CustomTestSetTitles Assert, TESTMODULE, "TestDropdownRejectsAnUnknownScope"
    If Not FixtureReady("TestDropdownRejectsAnUnknownScope") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As Linelist
    Dim drop As DropdownLists

    Set sut = Linelist.Create(Specs)
    Set drop = sut.Dropdown(7)

    CustomTestLogFailure Assert, "TestDropdownRejectsAnUnknownScope", , _
                         "Expected an error for a dropdown scope that does not exist"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Dropdown should raise for a scope a linelist has no sheet for"
End Sub

'@TestMethod("Linelist")
Public Sub TestPivotsKeepsTheInstanceItBuilt()
    CustomTestSetTitles Assert, TESTMODULE, "TestPivotsKeepsTheInstanceItBuilt"
    If Not FixtureReady("TestPivotsKeepsTheInstanceItBuilt") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As Linelist
    Dim firstRead As CustomPivotTable
    Dim secondRead As CustomPivotTable

    Set sut = Linelist.Create(Specs)
    Set firstRead = sut.Pivots
    Set secondRead = sut.Pivots

    Assert.IsTrue Not firstRead Is Nothing, _
                  "Pivots should build a manager over the custom pivot worksheet"
    Assert.IsTrue firstRead Is secondRead, _
                  "Pivots should answer the same manager, so the blocks it stacked stay known"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPivotsKeepsTheInstanceItBuilt", Err.Number, Err.Description
End Sub


'@section Lifecycle guard tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestPrepareRefusesUnpreparedSpecifications()
    CustomTestSetTitles Assert, TESTMODULE, "TestPrepareRefusesUnpreparedSpecifications"
    If Not FixtureReady("TestPrepareRefusesUnpreparedSpecifications") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As Linelist
    Set sut = Linelist.Create(Specs)
    sut.Prepare

    CustomTestLogFailure Assert, "TestPrepareRefusesUnpreparedSpecifications", , _
                         "Expected an error when the specifications have not been prepared"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number = ProjectError.ErrorUnexpectedState, _
                  "Prepare should raise ErrorUnexpectedState over unprepared specifications"
End Sub
