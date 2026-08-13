Attribute VB_Name = "TestSectionBuilder"
Attribute VB_Description = "Tests for SectionBuilder class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for SectionBuilder class")

Option Explicit

'@description
'Drives SectionBuilder, which groups the dictionary rows of one sheet into
'sections and subsections and hands each variable to a VarWriter.
'
'THE TWO SHEETS ARE BUILT ONCE FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'Building a sheet writes every label, every dropdown and every conditional
'format of that sheet, and character-level formatting is the slowest thing this
'project asks Excel for. A build per test put the runner past its cap with
'nothing else registered. So the fixture workbook, the VList sheet and the HList
'sheet are built once in ModuleInitialize and every test reads what they hold.
'The tests write nothing, so they cannot disturb each other.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize or TestInitialize reaches the VBE as a modal
'dialog and stops the whole run: no results file, and Excel left holding the
'staging copy. The setup captures its error instead and each test reports it as
'its own failure. This is the shape TestCodeTransfer uses.
'@depends SectionBuilder, VarWriter, LinelistSpecs, LLdictionary, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private Dict As LLdictionary
Private Specs As LinelistSpecs
Private TargetSheet As Worksheet
Private HListSheet As Worksheet
Private PrintSheet As Worksheet
Private DropStub As DropdownLists
Private CustDropStub As DropdownLists

'The two builders that laid the sheets out. They are held because what a
'build counted and recorded is read off the builder, and the sheets are
'built once for the whole module.
Private VListBuilder As SectionBuilder
Private HListBuilder As SectionBuilder

Private SetupError As Long
Private SetupMessage As String

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "SectionBuilder"
Private Const DICTIONARY_SHEET As String = "DictFixture"
Private Const CHOICES_SHEET As String = "ChoicesFixture"

'The data entry sheets carry the names the dictionary gives them, as the
'worksheets LLDataEntry creates do. LLSheets.VariableAddress drops the sheet
'prefix when the name matches.
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const HLIST_SHEET As String = "hlist2D-sheet1"

Private Const VLIST_SEC_COL As Long = 2
Private Const VLIST_SUBSEC_COL As Long = 3
Private Const VLIST_LABEL_COL As Long = 4
Private Const HLIST_SEC_ROW As Long = 5
Private Const HLIST_SUBSEC_ROW As Long = 6
Private Const HLIST_NAME_ROW As Long = 8


'@section Lifecycle
'===============================================================================

'@sub-title Build the fixture workbook and both data entry sheets, once.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestSectionBuilder"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        BuildFixture
        BuildBothSheets
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set TargetSheet = Nothing
    Set HListSheet = Nothing
    Set PrintSheet = Nothing
    Set DropStub = Nothing
    Set CustDropStub = Nothing
    Set VListBuilder = Nothing
    Set HListBuilder = Nothing
    Set FixtureWorkbook = Nothing
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


'@section Factory tests
'===============================================================================

'@TestMethod("SectionBuilder")
Public Sub TestCreateReturnsASectionBuilder()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsASectionBuilder"
    If Not FixtureReady("TestCreateReturnsASectionBuilder") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsASectionBuilder", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestSectionBuilderRaisesWithoutSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestSectionBuilderRaisesWithoutSpecs"
    If Not FixtureReady("TestSectionBuilderRaisesWithoutSpecs") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Nothing, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise when specs is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A missing specs object should raise ObjectNotInitialized - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestSectionBuilderRaisesWithoutWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestSectionBuilderRaisesWithoutWorksheet"
    If Not FixtureReady("TestSectionBuilderRaisesWithoutWorksheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=Nothing)

    Assert.LogFailure "Create should raise when wksh is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A missing worksheet should raise ObjectNotInitialized - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestSectionBuilderRejectsAnUnknownLayer()
    CustomTestSetTitles Assert, TESTMODULE, "TestSectionBuilderRejectsAnUnknownLayer"
    If Not FixtureReady("TestSectionBuilderRejectsAnUnknownLayer") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=7, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise for a layer that names neither member."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "A layer outside the enum should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@section What the VList build wrote
'===============================================================================

'@TestMethod("SectionBuilder")
Public Sub TestTheVListBuildWroteEverySection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListBuildWroteEverySection"
    If Not FixtureReady("TestTheVListBuildWroteEverySection") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue SectionIsWritten("Hidden Section"), "Section 'Hidden Section' should be written"
    Assert.IsTrue SectionIsWritten("Controls"), "Section 'Controls' should be written"
    Assert.IsTrue SectionIsWritten("Status"), "Section 'Status' should be written"
    Assert.IsTrue SectionIsWritten("Section only"), "Section 'Section only' should be written"
    Assert.IsTrue SectionIsWritten("Validation"), "Section 'Validation' should be written"
    Assert.IsTrue SectionIsWritten("Format"), "Section 'Format' should be written"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListBuildWroteEverySection", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheVListBuildWroteEverySubSection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListBuildWroteEverySubSection"
    If Not FixtureReady("TestTheVListBuildWroteEverySubSection") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue SubSectionIsWritten("Subsection only"), _
                  "Subsection 'Subsection only' should be written to column 3"
    Assert.IsTrue SubSectionIsWritten("Date validation"), _
                  "Subsection 'Date validation' should be written to column 3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListBuildWroteEverySubSection", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestASectionWithSeveralSubSectionsIsBuiltWhole()
    CustomTestSetTitles Assert, TESTMODULE, "TestASectionWithSeveralSubSectionsIsBuiltWhole"
    If Not FixtureReady("TestASectionWithSeveralSubSectionsIsBuiltWhole") Then Exit Sub
    On Error GoTo TestFail

    'The Validation section of vlist1D-sheet1 holds four subsections. A writer
    'was built per subsection and each of them keyed its checkings from 0, so
    'the merge raised on the second one and the sheet stopped there. The two
    'labels below sit after that section.
    Assert.IsTrue LabelIsWritten("Variable formated as currency"), _
                  "The build should reach the section that follows the four subsections"
    Assert.IsTrue LabelIsWritten("Hidden variable at the end"), _
                  "The build should reach the last variable of the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASectionWithSeveralSubSectionsIsBuiltWhole", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheVListBuildWroteTheVariableLabels()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListBuildWroteTheVariableLabels"
    If Not FixtureReady("TestTheVListBuildWroteTheVariableLabels") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue Application.WorksheetFunction.CountA(TargetSheet.Columns(VLIST_LABEL_COL)) > 0, _
                  "Variable labels should be written to column 4"
    Assert.IsTrue LabelIsWritten("Choices on vlist1D"), _
                  "The label of the first Controls variable should be written"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListBuildWroteTheVariableLabels", Err.Number, Err.Description
End Sub


'@section What the HList build wrote
'===============================================================================

'@TestMethod("SectionBuilder")
Public Sub TestTheHListBuildWroteTheSectionAndSubSectionRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListBuildWroteTheSectionAndSubSectionRows"
    If Not FixtureReady("TestTheHListBuildWroteTheSectionAndSubSectionRows") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue Application.WorksheetFunction.CountIf( _
                      HListSheet.Rows(HLIST_SEC_ROW), "Controls") > 0, _
                  "An HList section name should be written to row 5"
    Assert.IsTrue Application.WorksheetFunction.CountIf( _
                      HListSheet.Rows(HLIST_SUBSEC_ROW), "Subsection only") > 0, _
                  "An HList subsection name should be written to row 6"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListBuildWroteTheSectionAndSubSectionRows", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheHListBuildReachedTheLastVariable()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListBuildReachedTheLastVariable"
    If Not FixtureReady("TestTheHListBuildReachedTheLastVariable") Then Exit Sub
    On Error GoTo TestFail

    'The build was given no CRF companion, so every CRF lookup and every CRF
    'write is skipped. The crf index used to be read and converted above the
    'guard that tests for the companion.
    Assert.IsTrue Application.WorksheetFunction.CountIf( _
                      HListSheet.Rows(HLIST_NAME_ROW), "hid_end_h2") > 0, _
                  "A build with no CRF companion should reach the last variable of the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListBuildReachedTheLastVariable", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheHListBuildWroteThePrintCompanion()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListBuildWroteThePrintCompanion"
    If Not FixtureReady("TestTheHListBuildWroteThePrintCompanion") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue Application.WorksheetFunction.CountIf( _
                      PrintSheet.Rows(HLIST_SEC_ROW), "Controls") > 0, _
                  "The printed companion should carry the section headers too"
    Assert.IsTrue Application.WorksheetFunction.CountIf( _
                      PrintSheet.Rows(HLIST_NAME_ROW), "text_h2") > 0, _
                  "The printed companion should carry the variable names too"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListBuildWroteThePrintCompanion", Err.Number, Err.Description
End Sub


'@section What the build counted and recorded
'===============================================================================
'@description
'The build files one entry per section it lays out, and hands on the
'writer's per-variable record without merging it. The two travel to
'different places: the section entries reach the __check worksheet with
'the problems, the per-variable record reaches the run log's text file
'alone.

'@TestMethod("SectionBuilder")
Public Sub TestTheBuildFiledOneEntryPerSection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuildFiledOneEntryPerSection"
    If Not FixtureReady("TestTheBuildFiledOneEntryPerSection") Then Exit Sub
    On Error GoTo TestFail

    Dim map As SectionMap

    'The map holds one block per section the build laid out, so it is what
    'the count is measured against.
    Set map = SectionMap.Create(TargetSheet)

    Assert.IsTrue map.Count > 0, "The fixture sheet holds sections to place"
    Assert.AreEqual map.Count, VListBuilder.SectionsPlaced, _
                    "One section placed is one section counted"
    Assert.IsTrue VListBuilder.HasCheckings, _
                  "A build that placed a section should have something to report"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuildFiledOneEntryPerSection", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheSectionEntryNamesItsSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSectionEntryNamesItsSheet"
    If Not FixtureReady("TestTheSectionEntryNamesItsSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim entryKey As String

    entryKey = VLIST_SHEET & "!Controls"

    Assert.IsTrue VListBuilder.CheckingValues.KeyExists(entryKey), _
                  "The entry of a placed section is keyed by sheet and section"
    Assert.AreEqual "section Controls placed on " & VLIST_SHEET, _
                    VListBuilder.CheckingValues.ValueOf(entryKey, checkingLabel), _
                    "The entry says which section was placed on which sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSectionEntryNamesItsSheet", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheBuildCountsTheVariablesItWrote()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuildCountsTheVariablesItWrote"
    If Not FixtureReady("TestTheBuildCountsTheVariablesItWrote") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue VListBuilder.VariablesWritten > 0, _
                  "The VList build wrote variables and should count them"
    Assert.IsTrue HListBuilder.VariablesWritten > 0, _
                  "The HList build wrote variables and should count them"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuildCountsTheVariablesItWrote", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestTheMilestoneRecordIsTheWritersOwn()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMilestoneRecordIsTheWritersOwn"
    If Not FixtureReady("TestTheMilestoneRecordIsTheWritersOwn") Then Exit Sub
    On Error GoTo TestFail

    'The record is handed on rather than merged, so there is one entry per
    'variable written and no copy of it inside the section entries.
    Assert.IsTrue VListBuilder.HasMilestones, _
                  "A build that wrote variables carries their record"
    Assert.AreEqual CLng(VListBuilder.VariablesWritten), _
                    CLng(VListBuilder.MilestoneValues.Length), _
                    "One written variable is one entry of the record"
    Assert.IsTrue VListBuilder.CheckingValues.Length < VListBuilder.MilestoneValues.Length, _
                  "The per-variable record stays out of the entries the worksheet takes"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMilestoneRecordIsTheWritersOwn", Err.Number, Err.Description
End Sub


'@section What the build left in the section map
'===============================================================================
'@description
'Build records the boundaries of every section it lays out, so that a later
'session can act on a whole section without walking the dictionary again. These
'tests read the map back off the two sheets the fixture built.

'@TestMethod("SectionBuilder")
Public Sub TestTheVListBuildRecordedEverySectionRun()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListBuildRecordedEverySectionRun"
    If Not FixtureReady("TestTheVListBuildRecordedEverySectionRun") Then Exit Sub
    On Error GoTo TestFail

    Dim map As SectionMap
    Dim expectedNames As BetterArray
    Dim expectedStarts As BetterArray
    Dim expectedEnds As BetterArray
    Dim idx As Long

    'The expected runs are read back out of the dictionary here rather than
    'written down. `column index` is DERIVED by LLdictionary.Prepare and appears
    'in no fixture row, so a position written into this test would be a guess.
    LoadExpectedRuns VLIST_SHEET, expectedNames, expectedStarts, expectedEnds

    Set map = SectionMap.Create(TargetSheet)

    Assert.IsTrue expectedNames.Length > 0, _
                  "The fixture sheet holds sections to record"
    Assert.AreEqual CLng(expectedNames.Length), map.Count, _
                    "The build recorded one block per run of the dictionary"

    For idx = 1 To map.Count
        Assert.AreEqual CStr(expectedNames.Item(idx)), map.SectionNameAt(idx), _
                        "Block " & CStr(idx) & " carries the title of run " & CStr(idx)
        Assert.AreEqual CLng(expectedStarts.Item(idx)), map.StartAt(idx), _
                        "Block " & CStr(idx) & " starts where its first variable sits"
        Assert.AreEqual CLng(expectedEnds.Item(idx)), map.EndAt(idx), _
                        "Block " & CStr(idx) & " ends where its last variable sits"
    Next idx

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListBuildRecordedEverySectionRun", Err.Number, Err.Description
End Sub

'@TestMethod("SectionBuilder")
Public Sub TestThePreparedDictionaryGivesOneBlockPerTitle()
    CustomTestSetTitles Assert, TESTMODULE, "TestThePreparedDictionaryGivesOneBlockPerTitle"
    If Not FixtureReady("TestThePreparedDictionaryGivesOneBlockPerTitle") Then Exit Sub
    On Error GoTo TestFail

    Dim map As SectionMap
    Dim outer As Long
    Dim inner As Long
    Dim repeated As String

    'The fixture dictionary lists "Controls" in two places, once before the
    'Status variables and once after them. Those two runs still come out as one
    'block, because LLdictionary.Prepare sorts the sheet by main section before
    'it derives `column index`, and the sort brings the two runs together.
    '
    'So a title names one block on a prepared dictionary. SectionMap keys on the
    'position range all the same: the toggle starts from a selected cell, so a
    'position is what it has to look a section up by. TestSectionMap covers what
    'the map does when two blocks do share a title.
    Set map = SectionMap.Create(TargetSheet)

    Assert.IsTrue map.Count > 1, "The sheet carries several sections"

    For outer = 1 To map.Count
        For inner = outer + 1 To map.Count
            If LenB(map.SectionNameAt(outer)) > 0 Then
                If map.SectionNameAt(outer) = map.SectionNameAt(inner) Then
                    repeated = map.SectionNameAt(outer)
                End If
            End If
        Next inner
    Next outer

    Assert.AreEqual vbNullString, repeated, _
                    "The sort leaves each section title on one block only"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePreparedDictionaryGivesOneBlockPerTitle", Err.Number, Err.Description
End Sub

'@TestMethod("SectionBuilder")
Public Sub TestTheHListBuildRecordedItsSections()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListBuildRecordedItsSections"
    If Not FixtureReady("TestTheHListBuildRecordedItsSections") Then Exit Sub
    On Error GoTo TestFail

    Dim map As SectionMap

    Set map = SectionMap.Create(HListSheet)

    Assert.IsTrue map.Count > 0, _
                  "The HList build should have recorded its sections too"
    Assert.IsTrue BlocksAscendWithoutOverlapping(map), _
                  "The blocks run up the sheet in order and never overlap, " & _
                  "so every position belongs to at most one section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListBuildRecordedItsSections", Err.Number, Err.Description
End Sub

'@TestMethod("SectionBuilder")
Public Sub TestThePrintCompanionCarriesNoMap()
    CustomTestSetTitles Assert, TESTMODULE, "TestThePrintCompanionCarriesNoMap"
    If Not FixtureReady("TestThePrintCompanionCarriesNoMap") Then Exit Sub
    On Error GoTo TestFail

    'The map is written on the main sheet alone. The printed companion carries
    'the same columns and no section action is offered on it, so recording a
    'second copy there would be a second thing to keep in step for nothing.
    Assert.AreEqual CLng(0), SectionMap.Create(PrintSheet).Count, _
                    "The printed companion carries no section map"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePrintCompanionCarriesNoMap", Err.Number, Err.Description
End Sub


'@section Fixture
'===============================================================================

'@sub-title Build the workbook, the dictionary and the specifications.
Private Sub BuildFixture()
    Dim transStub As TranslationObject
    Dim design As LLFormat
    Dim formatSheet As Worksheet
    Dim formData As FormulaData
    Dim formulaSheet As Worksheet
    Dim choicesObj As LLChoices

    Set FixtureWorkbook = NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    ChoicesTestFixture.PrepareChoicesFixture CHOICES_SHEET, FixtureWorkbook
    Set choicesObj = LLChoices.Create(FixtureWorkbook.Worksheets(CHOICES_SHEET), 1, 1)

    'The three MSG_ tags AddLabel reads live in the table, the way a real
    'linelist translation table carries them.
    Set transStub = BuildTranslationObject(FixtureWorkbook, "ENG", _
        Array(Array("MSG_Calculated", "Calculated"), _
              Array("MSG_Mandatory", "Mandatory"), _
              Array("MSG_CustomChoice", "Custom choice")))
    Set formatSheet = LLFormatTestFixture.PrepareLLFormatFixture("LLFormatFixture", FixtureWorkbook)
    Set design = LLFormat.Create(formatSheet)
    Set formulaSheet = FormulaTestFixture.PrepareFormulaFixtureSheet("FormulaFixture", outwb:=FixtureWorkbook)
    Set formData = FormulaData.Create(formulaSheet)

    Set TargetSheet = FixtureWorkbook.Worksheets.Add
    TargetSheet.Name = VLIST_SHEET
    Set HListSheet = FixtureWorkbook.Worksheets.Add
    HListSheet.Name = HLIST_SHEET
    Set PrintSheet = FixtureWorkbook.Worksheets.Add

    'Dropdown lists live on their own host sheets (as in production) so their
    'list tables and validations never collide with the section content the
    'tests assert on TargetSheet.
    Set DropStub = DropdownLists.Create(FixtureWorkbook.Worksheets.Add, hprefix:=vbNullString)
    Set CustDropStub = DropdownLists.Create(FixtureWorkbook.Worksheets.Add, hprefix:=vbNullString)

    EnsureSpecsSheets FixtureWorkbook
    Set Specs = LinelistSpecs.Create(FixtureWorkbook)
    Specs.TestAssignDictionary Dict
    Specs.TestAssignDesignFormat design
    Specs.TestAssignTransObject transStub
    Specs.TestAssignFormulaData formData
    Specs.TestAssignChoices choicesObj
End Sub

'@sub-title Build the VList sheet and the HList sheet, once each.
Private Sub BuildBothSheets()
    Set VListBuilder = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
    VListBuilder.Build VLIST_SHEET, FindSheetStartRow(VLIST_SHEET)

    Set HListBuilder = SectionBuilder.Create( _
        layer:=SectionBuilderModeHList, _
        specs:=Specs, _
        wksh:=HListSheet, _
        printWksh:=PrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
    HListBuilder.Build HLIST_SHEET, FindSheetStartRow(HLIST_SHEET)
End Sub

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


'@section Helpers
'===============================================================================

'@description
'Read the section runs the dictionary describes for one sheet, in the order
'Build walks them: consecutive rows sharing a `main section` value are one run,
'and a run takes the `column index` of its first and last row.
'
'The expectations are read from the dictionary because `column index` is
'derived by LLdictionary.Prepare and appears in no fixture row. A position
'written into a test would be a number the test had guessed.
'@param sheetName String. The sheet to walk.
'@param names BetterArray. Set to the title of each run, in order.
'@param starts BetterArray. Set to the first position of each run.
'@param ends BetterArray. Set to the last position of each run.
Private Sub LoadExpectedRuns(ByVal sheetName As String, _
                             ByRef names As BetterArray, _
                             ByRef starts As BetterArray, _
                             ByRef ends As BetterArray)
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim indexRng As Range
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim runEnd As Long
    Dim runName As String
    Dim startPos As Long
    Dim endPos As Long

    Set names = New BetterArray
    Set starts = New BetterArray
    Set ends = New BetterArray
    names.LowerBound = 1
    starts.LowerBound = 1
    ends.LowerBound = 1

    Set sectionRng = Dict.DataRange("main section")
    Set sheetRng = Dict.DataRange("sheet name")
    Set indexRng = Dict.DataRange("column index")
    lastRow = Dict.Data.DataEndRow()

    rowIdx = FindSheetStartRow(sheetName)
    If rowIdx = 0 Then Exit Sub

    Do While rowIdx <= lastRow
        If CStr(sheetRng.Cells(rowIdx, 1).Value) <> sheetName Then Exit Do

        runName = CStr(sectionRng.Cells(rowIdx, 1).Value)
        runEnd = rowIdx

        Do While runEnd < lastRow
            If CStr(sectionRng.Cells(runEnd + 1, 1).Value) <> runName Then Exit Do
            If CStr(sheetRng.Cells(runEnd + 1, 1).Value) <> sheetName Then Exit Do
            runEnd = runEnd + 1
        Loop

        startPos = PositionAt(indexRng, rowIdx)
        endPos = PositionAt(indexRng, runEnd)

        'RecordSection skips a run whose bounds it cannot read, for the same
        'reason FormatSection leaves it alone.
        If startPos > 0 And endPos >= startPos Then
            names.Push runName
            starts.Push startPos
            ends.Push endPos
        End If

        rowIdx = runEnd + 1
    Loop
End Sub

'@description Read one `column index` cell as a whole number, answering 0 for
'anything unusable. The same rule SectionBuilder.NumberAt applies.
'@param indexRng Range. The `column index` column.
'@param rowIdx Long. The dictionary data row.
'@return Long. The stored position, or 0.
Private Function PositionAt(ByVal indexRng As Range, ByVal rowIdx As Long) As Long
    Dim cellValue As Variant

    cellValue = indexRng.Cells(rowIdx, 1).Value
    If IsError(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    If Not IsNumeric(cellValue) Then Exit Function

    PositionAt = CLng(cellValue)
End Function

'@description
'Test whether every block of a map starts after the one before it ends. A map
'whose blocks overlap would hand two sections the same position, and a rebuild
'that failed to clear the old map is one way that happens.
'@param map SectionMap. The map to walk.
'@return Boolean. True when the blocks ascend and never overlap.
Private Function BlocksAscendWithoutOverlapping(ByVal map As SectionMap) As Boolean
    Dim idx As Long

    For idx = 1 To map.Count
        If map.EndAt(idx) < map.StartAt(idx) Then Exit Function
        If idx > 1 Then
            If map.StartAt(idx) <= map.EndAt(idx - 1) Then Exit Function
        End If
    Next idx

    BlocksAscendWithoutOverlapping = True
End Function

'@description Test whether a section name reached the VList section column.
Private Function SectionIsWritten(ByVal sectionName As String) As Boolean
    SectionIsWritten = Application.WorksheetFunction.CountIf( _
        TargetSheet.Columns(VLIST_SEC_COL), sectionName) > 0
End Function

'@description Test whether a subsection name reached the VList subsection column.
Private Function SubSectionIsWritten(ByVal subSectionName As String) As Boolean
    SubSectionIsWritten = Application.WorksheetFunction.CountIf( _
        TargetSheet.Columns(VLIST_SUBSEC_COL), subSectionName) > 0
End Function

'@description Test whether a label opens one of the cells of the VList label column.
Private Function LabelIsWritten(ByVal mainLabel As String) As Boolean
    LabelIsWritten = Application.WorksheetFunction.CountIf( _
        TargetSheet.Columns(VLIST_LABEL_COL), mainLabel & "*") > 0
End Function

'@description Find the first row in the dictionary DataRange where sheet name matches.
Private Function FindSheetStartRow(ByVal sheetName As String) As Long
    Dim sheetRng As Range
    Dim endRow As Long
    Dim rowIdx As Long

    Set sheetRng = Dict.DataRange("sheet name")
    endRow = Dict.Data.DataEndRow()

    For rowIdx = 1 To endRow
        If CStr(sheetRng.Cells(rowIdx, 1).Value) = sheetName Then
            FindSheetStartRow = rowIdx
            Exit Function
        End If
    Next rowIdx

    FindSheetStartRow = 0
End Function
