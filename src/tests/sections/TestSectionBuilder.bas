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
    Dim vlistBuilder As SectionBuilder
    Dim hlistBuilder As SectionBuilder

    Set vlistBuilder = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
    vlistBuilder.Build VLIST_SHEET, FindSheetStartRow(VLIST_SHEET)

    Set hlistBuilder = SectionBuilder.Create( _
        layer:=SectionBuilderModeHList, _
        specs:=Specs, _
        wksh:=HListSheet, _
        printWksh:=PrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
    hlistBuilder.Build HLIST_SHEET, FindSheetStartRow(HLIST_SHEET)
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
