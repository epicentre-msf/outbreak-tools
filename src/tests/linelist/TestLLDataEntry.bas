Attribute VB_Name = "TestLLDataEntry"
Attribute VB_Description = "Tests for LLDataEntry class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLDataEntry class")

Option Explicit

'@description
'Drives LLDataEntry, which builds one data entry worksheet of a generated
'linelist. The suite it replaced had five tests and all five stopped at Create,
'so nothing below the factory had ever run.
'
'THE TWO SHEETS ARE BUILT ONCE FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'Building a data entry sheet writes every label, every dropdown and every
'conditional format of that sheet through SectionBuilder and VarWriter, and
'character-level formatting is the slowest thing this project asks Excel for. A
'build per test put the sections folder over the runner cap. So one workbook
'shaped like a designer and one standing in for the linelist are made once in
'ModuleInitialize, the VList sheet and the HList sheet are built once each, and
'every test reads what they hold.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'whole headless run comes back with no results file. The setup captures its
'error into two module fields and a guard at the top of every test reports it as
'that test's own failure. This is the shape TestCodeTransfer and
'TestSectionBuilder use.
'
'WHAT THE GOTO DROPDOWN TESTS WATCH
'-------------------------------------------------------------------------------
'The first row of vlist1D-sheet1 in the dictionary fixture is hid_beg_v1 and it
'is the only variable of its main section, "Hidden Section". CollectSectionNames
'used to index the dictionary one row lower than every other reader in the
'class, so the first variable's row was never examined and a section holding a
'single variable at the top of a sheet was left out of the dropdown with no
'message. TestTheVListGoToDropdownHoldsTheFirstSection is what states it.
'
'BOTH SHEETS ARRIVE CARRYING METADATA
'-------------------------------------------------------------------------------
'AddOutputSheet keeps a worksheet a template workbook already ships, so a build
'often runs over a sheet that already holds sheet_type, table_name,
'filtered_sheet and blank_row_count.
'HiddenNames.EnsureName writes its value only when it creates the name, so the
'fixture seeds all four before either build runs and every metadata assertion
'here reads a name the build had to overwrite.
'@depends LLDataEntry, Linelist, LinelistSpecs, LLdictionary, LLSheets,
'  LLFormat, LLChoices, FormulaData, DropdownLists, HiddenNames, CustomTest,
'  DictionaryTestFixture, ChoicesTestFixture, LLFormatTestFixture,
'  FormulaTestFixture, PasswordsTestFixture

Private Assert As CustomTest
Private SpecsWkb As Workbook
Private OutWkb As Workbook
Private Dict As LLdictionary
Private Specs As LinelistSpecs
Private LL As Linelist
Private SheetInfo As LLSheets
Private LongSpecs As LinelistSpecs
Private LongLL As Linelist
Private VListBuilder As LLDataEntry
Private HListBuilder As LLDataEntry
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLDataEntry"
Private Const DICTIONARY_SHEET As String = "DictFixture"
Private Const LONG_DICTIONARY_SHEET As String = "DictLongNames"
Private Const CHOICES_SHEET As String = "ChoicesFixture"

'The two data entry sheets carry the names the dictionary gives them, as the
'worksheets a generation creates do.
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const HLIST_SHEET As String = "hlist2D-sheet1"

'The sheets Dropdown and Pivots resolve. The translator the fixture installs
'answers every tag with the tag itself, so these are the names the class asks
'for.
Private Const SHEET_DROPDOWN_LISTS As String = "__dropdown_lists"
Private Const SHEET_CUSTOM_CHOICE As String = "LLSHEET_CustomChoice"
Private Const SHEET_CUSTOM_PIVOT As String = "LLSHEET_CustomPivotTable"

'The main section of the first dictionary row of vlist1D-sheet1, and the only
'variable it holds.
Private Const FIRST_VLIST_SECTION As String = "Hidden Section"

Private Const GOTO_LABEL As String = "Go to section"
Private Const GOTO_SECTION_CODE As String = "go_to_section"

'26 characters. An HList sheet allows 25, because Excel accepts 31 and the
'printed companion carries a six character prefix. A VList sheet has no
'companion, so the same name fits.
Private Const LONG_SHEET_NAME As String = "abcdefghijklmnopqrstuvwxyz"

'What LLDataEntry gives the printed and the filtered companion beyond their
'header row.
Private Const COMPANION_SPARE_ROWS As Long = 10

'What both data entry sheets carry as metadata before the build runs.
Private Const STALE_VALUE As String = "from the template"


'@section Lifecycle
'===============================================================================

'@sub-title Build both workbooks and both data entry sheets, once.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLDataEntry"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        BuildFixture
        BuildBothSheets
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print results and drop both workbooks.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    On Error Resume Next
        If Not SpecsWkb Is Nothing Then DeleteWorkbook SpecsWkb
        If Not OutWkb Is Nothing Then DeleteWorkbook OutWkb
    On Error GoTo 0

    Set VListBuilder = Nothing
    Set HListBuilder = Nothing
    Set LongLL = Nothing
    Set LongSpecs = Nothing
    Set SheetInfo = Nothing
    Set LL = Nothing
    Set Specs = Nothing
    Set Dict = Nothing
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
'under the default label.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Fixture
'===============================================================================

'@sub-title Build the designer-shaped workbook and the linelist workbook.
Private Sub BuildFixture()
    Dim transStub As TranslationObject
    Dim design As LLFormat
    Dim formatSheet As Worksheet
    Dim formData As FormulaData
    Dim formulaSheet As Worksheet
    Dim choicesObj As LLChoices
    Dim longDict As LLdictionary

    Set SpecsWkb = NewWorkbook()
    Set OutWkb = NewWorkbook()

    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, SpecsWkb
    Set Dict = LLdictionary.Create(SpecsWkb.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    ChoicesTestFixture.PrepareChoicesFixture CHOICES_SHEET, SpecsWkb
    Set choicesObj = LLChoices.Create(SpecsWkb.Worksheets(CHOICES_SHEET), 1, 1)

    'The tags AddLabel and the GoTo dropdown read. A tag the table does not
    'carry answers with the tag itself, which is what makes the three worksheet
    'names below resolve.
    Set transStub = BuildTranslationObject(SpecsWkb, "ENG", _
        Array(Array("MSG_Calculated", "Calculated"), _
              Array("MSG_Mandatory", "Mandatory"), _
              Array("MSG_CustomChoice", "Custom choice"), _
              Array("MSG_GoToSec", GOTO_LABEL)))

    Set formatSheet = LLFormatTestFixture.PrepareLLFormatFixture("LLFormatFixture", SpecsWkb)
    Set design = LLFormat.Create(formatSheet)
    Set formulaSheet = FormulaTestFixture.PrepareFormulaFixtureSheet("FormulaFixture", outwb:=SpecsWkb)
    Set formData = FormulaData.Create(formulaSheet)

    EnsureSpecsSheets SpecsWkb
    PasswordsTestFixture.PreparePasswordsFixture "__pass", SpecsWkb

    'The linelist workbook: the two data entry sheets a generation would have
    'added, and the three worksheets the cached managers read.
    EnsureWorksheet VLIST_SHEET, OutWkb, clearSheet:=True
    EnsureWorksheet HLIST_SHEET, OutWkb, clearSheet:=True
    EnsureWorksheet SHEET_DROPDOWN_LISTS, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
    EnsureWorksheet SHEET_CUSTOM_CHOICE, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
    EnsureWorksheet SHEET_CUSTOM_PIVOT, OutWkb, clearSheet:=True, visibility:=xlSheetHidden

    'Both data entry sheets go into the build already carrying metadata, which
    'is the template path: AddOutputSheet keeps a worksheet the template
    'workbook shipped. Every metadata assertion below reads a name that was
    'there before the build ran.
    SeedStaleMetadata OutWkb.Worksheets(VLIST_SHEET)
    SeedStaleMetadata OutWkb.Worksheets(HLIST_SHEET)

    Set Specs = LinelistSpecs.Create(SpecsWkb)
    Specs.TestAssignDictionary Dict
    Specs.TestAssignDesignFormat design
    Specs.TestAssignTransObject transStub
    Specs.TestAssignFormulaData formData
    Specs.TestAssignChoices choicesObj
    Specs.TestAssignLLWorkbook OutWkb

    Set LL = Linelist.Create(Specs)
    Set SheetInfo = LLSheets.Create(Dict)

    'A second dictionary whose every row belongs to one 26-character sheet
    'name, for the two length tests. It reaches the class through its own
    'specifications over the same designer workbook.
    Set longDict = BuildLongNameDictionary()
    Set LongSpecs = LinelistSpecs.Create(SpecsWkb)
    LongSpecs.TestAssignDictionary longDict
    LongSpecs.TestAssignDesignFormat design
    LongSpecs.TestAssignTransObject transStub
    LongSpecs.TestAssignFormulaData formData
    LongSpecs.TestAssignChoices choicesObj
    LongSpecs.TestAssignLLWorkbook OutWkb
    Set LongLL = Linelist.Create(LongSpecs)
End Sub

'@sub-title Give a worksheet the metadata a template sheet would arrive with.
'@details
'HiddenNames.EnsureName writes its value only when it creates the name, so a
'name that is already there keeps whatever it held. These four are what the
'build has to overwrite.
'@param sh Worksheet. The data entry sheet about to be built.
Private Sub SeedStaleMetadata(ByVal sh As Worksheet)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", STALE_VALUE, HiddenNameTypeString
    store.EnsureName "table_name", STALE_VALUE, HiddenNameTypeString
    store.EnsureName "filtered_sheet", STALE_VALUE, HiddenNameTypeString
    store.EnsureName "blank_row_count", 0, HiddenNameTypeLong
End Sub

'@sub-title A dictionary carrying one sheet name of 26 characters.
'@return LLdictionary. Prepared, over its own worksheet of the designer book.
Private Function BuildLongNameDictionary() As LLdictionary
    Dim sh As Worksheet
    Dim sheetNameColumn As Long
    Dim lastRow As Long
    Dim counter As Long
    Dim longDict As LLdictionary

    DictionaryTestFixture.PrepareDictionaryFixture LONG_DICTIONARY_SHEET, SpecsWkb
    Set sh = SpecsWkb.Worksheets(LONG_DICTIONARY_SHEET)

    sheetNameColumn = HeaderColumn(sh, "Sheet Name")
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row

    For counter = 2 To lastRow
        sh.Cells(counter, sheetNameColumn).Value = LONG_SHEET_NAME
    Next counter

    Set longDict = LLdictionary.Create(sh, 1, 1)
    longDict.Prepare
    Set BuildLongNameDictionary = longDict
End Function

'@sub-title The column of a header, by its text.
'@param sh Worksheet. The sheet carrying the header row.
'@param headerName String. The header to find.
'@return Long. The column index, or 0 when the header is absent.
Private Function HeaderColumn(ByVal sh As Worksheet, ByVal headerName As String) As Long
    Dim counter As Long
    Dim lastColumn As Long

    lastColumn = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column

    For counter = 1 To lastColumn
        If LCase$(Trim$(CStr(sh.Cells(1, counter).Value))) = LCase$(headerName) Then
            HeaderColumn = counter
            Exit Function
        End If
    Next counter
End Function

'@sub-title Build the VList sheet and the HList sheet, once each.
Private Sub BuildBothSheets()
    Set VListBuilder = LLDataEntry.Create(LLDataEntryLayerVList, VLIST_SHEET, LL, SheetInfo)
    VListBuilder.Build

    Set HListBuilder = LLDataEntry.Create(LLDataEntryLayerHList, HLIST_SHEET, LL, SheetInfo)
    HListBuilder.Build
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

'@fun-title The table name the dictionary gives a sheet.
'@param sheetName String. The dictionary sheet name.
'@return String. The sheet's table name.
Private Function TableNameOf(ByVal sheetName As String) As String
    TableNameOf = SheetInfo.SheetInfo(sheetName, SheetInfoSheetTable)
End Function

'@fun-title Whether a BetterArray holds an entry.
'@param list BetterArray. The list to walk.
'@param wanted String. The entry looked for.
'@return Boolean. True when the entry is there.
Private Function ListHolds(ByVal list As BetterArray, ByVal wanted As String) As Boolean
    Dim counter As Long

    If list Is Nothing Then Exit Function
    If list.Length = 0 Then Exit Function

    For counter = list.LowerBound To list.UpperBound
        If CStr(list.Item(counter)) = wanted Then
            ListHolds = True
            Exit Function
        End If
    Next counter
End Function


'@section Factory tests
'===============================================================================

'@TestMethod("LLDataEntry")
Public Sub TestCreateHListReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateHListReturnsInstance"
    If Not FixtureReady("TestCreateHListReturnsInstance") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerHList, HLIST_SHEET, LL, SheetInfo)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance for the HList layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateHListReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestCreateVListReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateVListReturnsInstance"
    If Not FixtureReady("TestCreateVListReturnsInstance") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerVList, VLIST_SHEET, LL, SheetInfo)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance for the VList layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateVListReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestCreateRejectsNothingLinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingLinelist"
    On Error GoTo ExpectError

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerHList, "SomeSheet", Nothing)

    Assert.LogFailure "Create should raise when the linelist is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A Nothing linelist should raise ObjectNotInitialized - " & _
                    "description was [" & Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestCreateRejectsEmptySheetName()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsEmptySheetName"
    If Not FixtureReady("TestCreateRejectsEmptySheetName") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerHList, vbNullString, LL, SheetInfo)

    Assert.LogFailure "Create should raise when the sheet name is empty."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "An empty sheet name should raise InvalidArgument - " & _
                    "description was [" & Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestCreateRejectsUnknownSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsUnknownSheet"
    If Not FixtureReady("TestCreateRejectsUnknownSheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerHList, "NonExistentSheet__xyz", LL, SheetInfo)

    Assert.LogFailure "Create should raise when the sheet is absent from the dictionary."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "A sheet absent from the dictionary should raise InvalidArgument - " & _
                    "description was [" & Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestCreateRejectsAnUnknownLayer()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsAnUnknownLayer"
    If Not FixtureReady("TestCreateRejectsAnUnknownLayer") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(7, VLIST_SHEET, LL, SheetInfo)

    Assert.LogFailure "Create should raise for a layer that is neither HList nor VList."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "An unknown layer should be refused by Create - " & _
                    "description was [" & Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheInstanceRefusesASecondSheetName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheInstanceRefusesASecondSheetName"
    If Not FixtureReady("TestTheInstanceRefusesASecondSheetName") Then Exit Sub

    Dim sut As LLDataEntry
    On Error GoTo TestFail
    Set sut = LLDataEntry.Create(LLDataEntryLayerVList, VLIST_SHEET, LL, SheetInfo)

    On Error GoTo ExpectError
    sut.BinSheetName = "AnotherSheet"

    Assert.LogFailure "A sealed instance should refuse a second sheet name."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), CLng(Err.Number), _
                    "A sealed instance should refuse a creation-only write - " & _
                    "description was [" & Err.Description & "]"
    Err.Clear
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestTheInstanceRefusesASecondSheetName", Err.Number, Err.Description
End Sub


'@section Sheet name length
'===============================================================================

'@TestMethod("LLDataEntry")
Public Sub TestALongHListSheetNameIsReportedAtCreate()
    CustomTestSetTitles Assert, TESTMODULE, "TestALongHListSheetNameIsReportedAtCreate"
    If Not FixtureReady("TestALongHListSheetNameIsReportedAtCreate") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerHList, LONG_SHEET_NAME, LongLL)

    Assert.IsTrue sut.HasCheckings, _
                  "A 26-character HList sheet name should be reported at Create, " & _
                  "because its printed companion is cut to 31 characters"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongHListSheetNameIsReportedAtCreate", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheSameNameFitsAVListSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSameNameFitsAVListSheet"
    If Not FixtureReady("TestTheSameNameFitsAVListSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As LLDataEntry
    Set sut = LLDataEntry.Create(LLDataEntryLayerVList, LONG_SHEET_NAME, LongLL)

    Assert.IsTrue Not sut.HasCheckings, _
                  "A VList sheet has no prefixed companion, so 26 characters fit " & _
                  "and nothing is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSameNameFitsAVListSheet", Err.Number, Err.Description
End Sub


'@section VList build
'===============================================================================

'@TestMethod("LLDataEntry")
Public Sub TestAVListSheetCarriesItsType()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListSheetCarriesItsType"
    If Not FixtureReady("TestAVListSheetCarriesItsType") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Set store = HiddenNames.Create(OutWkb.Worksheets(VLIST_SHEET))

    Assert.AreEqual "VList", store.ValueAsString("sheet_type"), _
                    "A built VList sheet should record what kind of sheet it is"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListSheetCarriesItsType", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAVListSheetCarriesItsTableName()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListSheetCarriesItsTableName"
    If Not FixtureReady("TestAVListSheetCarriesItsTableName") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Set store = HiddenNames.Create(OutWkb.Worksheets(VLIST_SHEET))

    Assert.AreEqual TableNameOf(VLIST_SHEET), store.ValueAsString("table_name"), _
                    "A built VList sheet should record the table name the dictionary gave it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListSheetCarriesItsTableName", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAVListSheetCarriesItsGoToLabelInA1()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListSheetCarriesItsGoToLabelInA1"
    If Not FixtureReady("TestAVListSheetCarriesItsGoToLabelInA1") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = OutWkb.Worksheets(VLIST_SHEET)

    Assert.AreEqual GOTO_LABEL, CStr(sh.Cells(1, 1).Value), _
                    "Cell A1 of a data entry sheet carries the GoTo section label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListSheetCarriesItsGoToLabelInA1", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheVListGoToDropdownHoldsTheFirstSection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListGoToDropdownHoldsTheFirstSection"
    If Not FixtureReady("TestTheVListGoToDropdownHoldsTheFirstSection") Then Exit Sub
    On Error GoTo TestFail

    Dim entries As BetterArray
    Dim listName As String

    listName = TableNameOf(VLIST_SHEET) & "_" & GOTO_SECTION_CODE
    Set entries = LL.Dropdown(1).Values(listName)

    Assert.IsTrue entries.Length > 0, _
                  "The GoTo dropdown of the VList sheet should hold its sections"
    Assert.IsTrue ListHolds(entries, GOTO_LABEL & ": " & FIRST_VLIST_SECTION), _
                  "The first dictionary row of the sheet is the only variable of " & _
                  FIRST_VLIST_SECTION & ", and that section belongs in the dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListGoToDropdownHoldsTheFirstSection", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheVListGoToDropdownKeepsEachSectionOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListGoToDropdownKeepsEachSectionOnce"
    If Not FixtureReady("TestTheVListGoToDropdownKeepsEachSectionOnce") Then Exit Sub
    On Error GoTo TestFail

    Dim entries As BetterArray
    Dim seen As Collection
    Dim counter As Long
    Dim entry As String
    Dim duplicates As Long

    Set entries = LL.Dropdown(1).Values(TableNameOf(VLIST_SHEET) & "_" & GOTO_SECTION_CODE)
    Set seen = New Collection

    For counter = entries.LowerBound To entries.UpperBound
        entry = CStr(entries.Item(counter))
        On Error Resume Next
            seen.Add entry, entry
            If Err.Number <> 0 Then duplicates = duplicates + 1
            Err.Clear
        On Error GoTo TestFail
    Next counter

    Assert.AreEqual CLng(0), duplicates, _
                    "A section that repeats on consecutive dictionary rows is kept once"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListGoToDropdownKeepsEachSectionOnce", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAVListSheetNamesItsValuesRange()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListSheetNamesItsValuesRange"
    If Not FixtureReady("TestAVListSheetNamesItsValuesRange") Then Exit Sub
    On Error GoTo TestFail

    Dim rng As Range
    Dim sh As Worksheet

    Set sh = OutWkb.Worksheets(VLIST_SHEET)

    On Error Resume Next
        Set rng = sh.Range(TableNameOf(VLIST_SHEET) & "_PLAGEVALUES")
    On Error GoTo TestFail

    Assert.IsTrue Not rng Is Nothing, _
                  "A built VList sheet carries a PLAGEVALUES range over its value cells"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListSheetNamesItsValuesRange", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAVListSheetIsProtectedWhenTheBuildEnds()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListSheetIsProtectedWhenTheBuildEnds"
    If Not FixtureReady("TestAVListSheetIsProtectedWhenTheBuildEnds") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue OutWkb.Worksheets(VLIST_SHEET).ProtectContents, _
                  "The last statement of a build protects the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListSheetIsProtectedWhenTheBuildEnds", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheVListBuildFilesAReportEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheVListBuildFilesAReportEntry"
    If Not FixtureReady("TestTheVListBuildFilesAReportEntry") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue VListBuilder.HasCheckings, _
                  "A build files at least the line saying the sheet was built"
    Assert.IsTrue Not VListBuilder.CheckingValues Is Nothing, _
                  "CheckingValues answers the entries once there are some"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheVListBuildFilesAReportEntry", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheBuildOverwritesMetadataATemplateSheetCarried()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuildOverwritesMetadataATemplateSheetCarried"
    If Not FixtureReady("TestTheBuildOverwritesMetadataATemplateSheetCarried") Then Exit Sub
    On Error GoTo TestFail

    Dim vlistStore As HiddenNames
    Dim hlistStore As HiddenNames

    Set vlistStore = HiddenNames.Create(OutWkb.Worksheets(VLIST_SHEET))
    Set hlistStore = HiddenNames.Create(OutWkb.Worksheets(HLIST_SHEET))

    'All four names were on the two worksheets before either build ran, which
    'is what AddOutputSheet hands the class on the template path. EnsureName
    'writes a value only when it creates the name, so the build has to set them.
    Assert.IsTrue vlistStore.ValueAsString("sheet_type") <> STALE_VALUE, _
                  "The VList build overwrites the sheet type the template carried"
    Assert.IsTrue vlistStore.ValueAsString("table_name") <> STALE_VALUE, _
                  "The VList build overwrites the table name the template carried"
    Assert.IsTrue hlistStore.ValueAsString("filtered_sheet") <> STALE_VALUE, _
                  "The HList build overwrites the filtered sheet the template carried"
    Assert.IsTrue hlistStore.ValueAsLong("blank_row_count") > 0, _
                  "The HList build overwrites the blank row count the template carried"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuildOverwritesMetadataATemplateSheetCarried", Err.Number, Err.Description
End Sub


'@section HList build
'===============================================================================

'@TestMethod("LLDataEntry")
Public Sub TestAnHListSheetCarriesItsType()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnHListSheetCarriesItsType"
    If Not FixtureReady("TestAnHListSheetCarriesItsType") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Set store = HiddenNames.Create(OutWkb.Worksheets(HLIST_SHEET))

    Assert.AreEqual "HList", store.ValueAsString("sheet_type"), _
                    "A built HList sheet should record what kind of sheet it is"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnHListSheetCarriesItsType", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAnHListSheetCarriesItsTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnHListSheetCarriesItsTable"
    If Not FixtureReady("TestAnHListSheetCarriesItsTable") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tableName As String
    Dim found As Boolean

    Set sh = OutWkb.Worksheets(HLIST_SHEET)
    tableName = TableNameOf(HLIST_SHEET)

    On Error Resume Next
        found = (sh.ListObjects(tableName).Name = tableName)
    On Error GoTo TestFail

    Assert.IsTrue found, _
                  "A built HList sheet carries a ListObject under its dictionary table name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnHListSheetCarriesItsTable", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestThePrintCompanionCarriesItsOwnType()
    CustomTestSetTitles Assert, TESTMODULE, "TestThePrintCompanionCarriesItsOwnType"
    If Not FixtureReady("TestThePrintCompanionCarriesItsOwnType") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Set store = HiddenNames.Create(OutWkb.Worksheets("print_" & HLIST_SHEET))

    Assert.AreEqual "HList Print", store.ValueAsString("sheet_type"), _
                    "The printed companion records itself as a print sheet"
    Assert.AreEqual "print_" & TableNameOf(HLIST_SHEET), store.ValueAsString("table_name"), _
                    "The printed companion records its own table name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePrintCompanionCarriesItsOwnType", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheFilteredCompanionIsNamedOnTheDataSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFilteredCompanionIsNamedOnTheDataSheet"
    If Not FixtureReady("TestTheFilteredCompanionIsNamedOnTheDataSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Dim filteredName As String

    Set store = HiddenNames.Create(OutWkb.Worksheets(HLIST_SHEET))
    filteredName = store.ValueAsString("filtered_sheet")

    Assert.AreEqual "f" & HLIST_SHEET, filteredName, _
                    "The data sheet records where its filtered companion is"
    Assert.IsTrue LL.SheetExists(filteredName), _
                  "The filtered companion is a worksheet of the output workbook"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFilteredCompanionIsNamedOnTheDataSheet", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheFilteredCompanionTakesTheHeaderAndRoomToGrow()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFilteredCompanionTakesTheHeaderAndRoomToGrow"
    If Not FixtureReady("TestTheFilteredCompanionTakesTheHeaderAndRoomToGrow") Then Exit Sub
    On Error GoTo TestFail

    Dim sourceTable As ListObject
    Dim filteredTable As ListObject

    Set sourceTable = OutWkb.Worksheets(HLIST_SHEET).ListObjects(TableNameOf(HLIST_SHEET))
    Set filteredTable = OutWkb.Worksheets("f" & HLIST_SHEET).ListObjects("f" & TableNameOf(HLIST_SHEET))

    Assert.AreEqual CStr(sourceTable.HeaderRowRange.Cells(1, 1).Value), _
                    CStr(filteredTable.HeaderRowRange.Cells(1, 1).Value), _
                    "The filtered companion takes the header row of the source table"
    Assert.AreEqual CLng(COMPANION_SPARE_ROWS + 1), _
                    CLng(filteredTable.ListRows.Count), _
                    "The filtered companion is given the header row and room to grow"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFilteredCompanionTakesTheHeaderAndRoomToGrow", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheBlankRowCountIsWhatAnUntouchedRowCarries()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBlankRowCountIsWhatAnUntouchedRowCarries"
    If Not FixtureReady("TestTheBlankRowCountIsWhatAnUntouchedRowCarries") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Dim sourceTable As ListObject
    Dim filled As Long

    Set store = HiddenNames.Create(OutWkb.Worksheets(HLIST_SHEET))
    Set sourceTable = OutWkb.Worksheets(HLIST_SHEET).ListObjects(TableNameOf(HLIST_SHEET))
    filled = Application.WorksheetFunction.CountA(sourceTable.ListRows(1).Range)

    Assert.AreEqual CLng(filled), CLng(store.ValueAsLong("blank_row_count")), _
                    "blank_row_count holds the filled cells of an untouched data row, " & _
                    "which is the threshold LLImporter and the buttons compare against"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBlankRowCountIsWhatAnUntouchedRowCarries", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestAnHListSheetIsProtectedWhenTheBuildEnds()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnHListSheetIsProtectedWhenTheBuildEnds"
    If Not FixtureReady("TestAnHListSheetIsProtectedWhenTheBuildEnds") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue OutWkb.Worksheets(HLIST_SHEET).ProtectContents, _
                  "The last statement of a build protects the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnHListSheetIsProtectedWhenTheBuildEnds", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheHListGoToDropdownHoldsItsSections()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListGoToDropdownHoldsItsSections"
    If Not FixtureReady("TestTheHListGoToDropdownHoldsItsSections") Then Exit Sub
    On Error GoTo TestFail

    Dim entries As BetterArray
    Set entries = LL.Dropdown(1).Values(TableNameOf(HLIST_SHEET) & "_" & GOTO_SECTION_CODE)

    Assert.IsTrue entries.Length > 0, _
                  "The GoTo dropdown of the HList sheet should hold its sections"
    Assert.IsTrue ListHolds(entries, GOTO_LABEL & ": Controls"), _
                  "Controls is a main section of hlist2D-sheet1 and belongs in the dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListGoToDropdownHoldsItsSections", Err.Number, Err.Description
End Sub

'@TestMethod("LLDataEntry")
Public Sub TestTheHListBuildFilesAReportEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHListBuildFilesAReportEntry"
    If Not FixtureReady("TestTheHListBuildFilesAReportEntry") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue HListBuilder.HasCheckings, _
                  "A build files at least the line saying the sheet was built"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHListBuildFilesAReportEntry", Err.Number, Err.Description
End Sub


'@section What the build counted
'===============================================================================
'@description
'The line a finished sheet files carries what the sheet holds, and the driver
'adds the counts up over the run for the closing bundle of the report.

'@TestMethod("LLDataEntry")
Public Sub TestTheSheetEntryCarriesItsCounts()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSheetEntryCarriesItsCounts"
    If Not FixtureReady("TestTheSheetEntryCarriesItsCounts") Then Exit Sub
    On Error GoTo TestFail

    Dim entries As Checking
    Dim keys As BetterArray
    Dim idx As Long
    Dim joined As String

    Set entries = VListBuilder.CheckingValues
    Set keys = entries.ListOfKeys()

    For idx = keys.LowerBound To keys.UpperBound
        joined = joined & entries.ValueOf(CStr(keys.Item(idx)), checkingLabel) & " | "
    Next idx

    Assert.IsTrue InStr(1, joined, "VList sheet " & VLIST_SHEET & " built") > 0, _
                  "The finished sheet files its line, and the entries read " & joined
    Assert.IsTrue InStr(1, joined, "section(s), ") > 0, _
                  "The line carries the section count, and the entries read " & joined
    Assert.IsTrue InStr(1, joined, CStr(VListBuilder.VariablesWritten) & " variable(s)") > 0, _
                  "The line carries the variable count, and the entries read " & joined

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSheetEntryCarriesItsCounts", Err.Number, Err.Description
End Sub


'@TestMethod("LLDataEntry")
Public Sub TestTheBuildHandsOnThePerVariableRecord()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuildHandsOnThePerVariableRecord"
    If Not FixtureReady("TestTheBuildHandsOnThePerVariableRecord") Then Exit Sub
    On Error GoTo TestFail

    'The driver takes this store record-only, so it reaches the run's text
    'file and stays off the __check worksheet.
    Assert.IsTrue HListBuilder.HasMilestones, _
                  "A sheet that wrote variables hands on their record"
    Assert.IsTrue HListBuilder.VariablesWritten > 0, _
                  "A built sheet counts the variables it wrote"
    Assert.AreEqual CLng(HListBuilder.VariablesWritten), _
                    CLng(HListBuilder.MilestoneValues.Length), _
                    "One written variable is one entry of the record"
    Assert.IsTrue HListBuilder.SectionsPlaced > 0, _
                  "A built sheet counts the sections it placed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuildHandsOnThePerVariableRecord", Err.Number, Err.Description
End Sub
