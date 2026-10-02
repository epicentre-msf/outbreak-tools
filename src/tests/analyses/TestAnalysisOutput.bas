Attribute VB_Name = "TestAnalysisOutput"
Attribute VB_Description = "Tests for AnalysisOutput class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnalysisOutput class")

'@description
'Drives AnalysisOutput, the class that walks the analysis setup tables and
'writes one analysis worksheet per scope. What this suite watches is the
'SECTION LOOP: which section each drawn table is filed under, and how many
'entries each section's navigation dropdown ends up with.
'
'WHY THE ASSERTIONS ARE AT THIS LEVEL
'-------------------------------------------------------------------------------
'Issue 337 reported a section dropdown listing one entry fewer than the setup
'defines. The defect was in the loop that carries the open section, and
'TemporalSection in isolation was never wrong, so a test of that class could
'never have caught it. These tests read the dropdown the loop produced.
'
'THE FIXTURE IS TWO WORKBOOKS, BUILT ONCE PER MODULE
'-------------------------------------------------------------------------------
'One workbook stands in for the designer and carries the dictionary, the
'choices, the format table, the formula tokens, the seven worksheets
'LinelistSpecs.Create validates, and the analysis setup ListObject. The other
'stands in for the linelist and carries the four analysis output worksheets plus
'the internal sheets the class resolves. The second reaches LinelistSpecs
'through TestAssignLLWorkbook, the hook Prepare writes in production.
'
'The translator is a pass-through, so every tag answers with its own name. That
'is why the output worksheets are called LLSHEET_Analysis and the rest: those
'are the tags AnalysisOutput asks the translator for.
'
'The designer-shaped workbook is built once. The linelist one is remade per
'build, because a build protects its sheets and leaves names behind, and a new
'workbook carries none of that.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'whole headless run comes back with no results file. The setup captures the
'error into two module fields and a guard at the top of every test reports it as
'that test's own failure.
'
'WHAT THIS SUITE CANNOT REACH
'-------------------------------------------------------------------------------
'Issue 337 has a second arrangement: a row that passes validation and then
'raises inside Build. Making Excel refuse a write on one chosen row, and only
'that row, needs the sheet in a state no fixture can put it in from here. What
'the suite reaches instead is the arrangement that shares the observable -- a
'section whose anchor row never becomes an anchor -- through a row the
'validation rejects.
'@depends AnalysisOutput, Linelist, LinelistSpecs, LLdictionary, LLChoices,
'  LLFormat, FormulaData, TranslationObject, CustomTest, TestHelpersLite,
'  DictionaryTestFixture, LLFormatTestFixture, FormulaTestFixture, ChoicesTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

' Sheets of the workbook standing in for the designer.
Private Const DICT_SHEET As String = "AODictFixture"
Private Const CHOICES_SHEET As String = "AOChoicesFixture"
Private Const FORMAT_SHEET As String = "AOFormatFixture"
Private Const TOKENS_SHEET As String = "AOTokensFixture"
Private Const SPECS_SHEET As String = "AOAnalysisSpecs"

' Sheets of the workbook standing in for the linelist. The translator is a
' pass-through, so a sheet is named after the tag the class asks for.
Private Const SHEET_NORMAL As String = "LLSHEET_Analysis"
Private Const SHEET_TEMPORAL As String = "LLSHEET_TemporalAnalysis"
Private Const SHEET_SPATIAL As String = "LLSHEET_SpatialAnalysis"
Private Const SHEET_SPATIOTEMPORAL As String = "LLSHEET_SpatioTemporalAnalysis"
Private Const SHEET_CUSTOM_CHOICE As String = "LLSHEET_CustomChoice"
Private Const SHEET_CUSTOM_PIVOT As String = "LLSHEET_CustomPivotTable"
Private Const SHEET_ANA_NAMES As String = "__ana_tabnames"
Private Const SHEET_DROPDOWN_LISTS As String = "__dropdown_lists"
Private Const SHEET_SPATIAL_TABLES As String = "__spatial_tables"

' The analysis setup ListObject this suite writes, spelled the way the setup
' workbook spells it.
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"

' The header row of the setup table, and where it sits on the specs sheet.
Private Const HEADER_ROW As Long = 3

' Variables of the shared dictionary fixture.
Private Const DATE_VARIABLE As String = "date_v1"
Private Const COL_CHOICE_VARIABLE As String = "choi_ord_v1"

' A variable name the dictionary does not carry, which is what makes a setup row
' fail validation.
Private Const UNKNOWN_VARIABLE As String = "no_such_variable"

' The two section labels the fixture uses, and the prefix AnalysisOutput gives
' the navigation dropdowns of the time series sheet.
Private Const SECTION_ONE As String = "First section"
Private Const SECTION_TWO As String = "Second section"
Private Const TEMPORAL_GOTO_PREFIX As String = "ts_"

Private Assert As CustomTest
Private SpecsWkb As Workbook
Private OutWkb As Workbook
Private Specs As LinelistSpecs
Private Dict As LLdictionary
Private LL As Linelist
Private SetupError As Long
Private SetupMessage As String


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
    Assert.SetModuleName "TestAnalysisOutput"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set SpecsWkb = NewWorkbook()

        PrepareDictionaryFixture DICT_SHEET, SpecsWkb
        Set Dict = LLdictionary.Create(SpecsWkb.Worksheets(DICT_SHEET), 1, 1)
        Dict.Prepare

        PrepareChoicesFixture CHOICES_SHEET, SpecsWkb
        EnsureSpecsSheets SpecsWkb

        ' The password sheet EnsureSpecsSheets leaves empty. LinelistSpecs
        ' builds a Passwords over it and BuildScope protects every sheet it
        ' finishes, so the named ranges and the two tables have to be there.
        PreparePasswordsFixture "__pass", SpecsWkb

        Set Specs = LinelistSpecs.Create(SpecsWkb)
        Specs.TestAssignDictionary Dict
        Specs.TestAssignTransObject BuildTranslationObject(SpecsWkb, "ENG", Array())
        Specs.TestAssignChoices LLChoices.Create(SpecsWkb.Worksheets(CHOICES_SHEET), 1, 1)
        Specs.TestAssignDesignFormat LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET, SpecsWkb))
        ' FormulaData resolves its two lookup tables by FIXED name, so the
        ' fixture has to use those names. They sit in the workbook standing in
        ' for the designer, and a ListObject name is unique per workbook, so the
        ' copies the other suites build in this one are untouched.
        Specs.TestAssignFormulaData FormulaData.Create( _
            PrepareFormulaFixtureSheet(TOKENS_SHEET, outwb:=SpecsWkb))

        NewLinelistWorkbook

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print results and drop the two workbooks.
'@details
'This routine is Public because the harness reaches it through Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    ' The freeze-pane pass activates worksheets of the linelist workbook, so the
    ' screen has to be handed back before the harness writes its results into a
    ' worksheet of THIS workbook.
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    On Error Resume Next
        If Not SpecsWkb Is Nothing Then DeleteWorkbook SpecsWkb
        If Not OutWkb Is Nothing Then DeleteWorkbook OutWkb
    On Error GoTo 0

    Set LL = Nothing
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
'under the default label.
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

'@sub-title The header row of Tab_TimeSeries_Analysis, 12 columns.
'@details
'Measured in the setup workbook and shared with TestCrossTable by copy, because
'a suite keeps what it needs inside itself.
Private Function TimeSeriesHeader() As Variant
    TimeSeriesHeader = Array( _
        "Series ID", "Section", "Time variable (row)", _
        "Group by variable (column)", "Title (header)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add total", "Table order")
End Function

'@sub-title One time series setup row.
'@param sectionLabel String. The section the row belongs to.
'@param timeVar String. The time variable of the row.
'@param title String. The table title.
'@param tableOrder String. The order column.
'@return Variant. One row array as wide as the header.
Private Function TimeSeriesRow(ByVal sectionLabel As String, _
                               ByVal timeVar As String, _
                               ByVal title As String, _
                               ByVal tableOrder As String) As Variant
    TimeSeriesRow = Array("Series " & tableOrder, sectionLabel, timeVar, _
                          COL_CHOICE_VARIABLE, title, "no", "", "Cases", _
                          "", "", "no", tableOrder)
End Function

'@sub-title Two sections of three tables each, every row valid.
Private Function TwoWholeSections() As Variant
    TwoWholeSections = Array( _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First A", "1"), _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First B", "2"), _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First C", "3"), _
        TimeSeriesRow(SECTION_TWO, DATE_VARIABLE, "Second A", "4"), _
        TimeSeriesRow(SECTION_TWO, DATE_VARIABLE, "Second B", "5"), _
        TimeSeriesRow(SECTION_TWO, DATE_VARIABLE, "Second C", "6"))
End Function

'@sub-title Two sections where the anchor row of the second one is rejected.
'@details
'The anchor names a time variable the dictionary does not carry, so ValidTable
'answers False and the row is never drawn. What the section below it must not
'do is append its tables to the section above.
Private Function SecondSectionWithABadAnchor() As Variant
    SecondSectionWithABadAnchor = Array( _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First A", "1"), _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First B", "2"), _
        TimeSeriesRow(SECTION_ONE, DATE_VARIABLE, "First C", "3"), _
        TimeSeriesRow(SECTION_TWO, UNKNOWN_VARIABLE, "Second A", "4"), _
        TimeSeriesRow(SECTION_TWO, DATE_VARIABLE, "Second B", "5"), _
        TimeSeriesRow(SECTION_TWO, DATE_VARIABLE, "Second C", "6"))
End Function

'@sub-title Write the analysis setup ListObject on the specifications workbook.
'@param dataRows Variant. An array of row arrays, each as wide as the header.
'@return Worksheet. The specifications worksheet.
Private Function BuildSetupTable(ByVal dataRows As Variant) As Worksheet
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tableRng As Range
    Dim headerRow As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim idx As Long

    headerRow = TimeSeriesHeader()
    colCount = UBound(headerRow) - LBound(headerRow) + 1
    rowCount = UBound(dataRows) - LBound(dataRows) + 1

    Set sh = EnsureWorksheet(SPECS_SHEET, SpecsWkb, clearSheet:=False, _
                             visibility:=xlSheetVisible)

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx
    sh.Cells.Clear

    WriteMatrix sh.Cells(HEADER_ROW, 1), RowsToMatrix(Array(headerRow))
    WriteMatrix sh.Cells(HEADER_ROW + 1, 1), RowsToMatrix(dataRows)

    Set tableRng = sh.Range(sh.Cells(HEADER_ROW, 1), _
                            sh.Cells(HEADER_ROW + rowCount, colCount))
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRng, , xlYes)
    lo.Name = TABLE_TIMESERIES

    Set BuildSetupTable = sh
End Function

'@sub-title Make a fresh workbook to stand in for the linelist.
'@details
'One per build. A build protects every worksheet it finishes, leaves
'workbook-scoped names behind that a clear does not remove, and grows the
'dropdown registry on a worksheet of the workbook. Putting all of that back by
'hand needs the password the build protected with, and it takes more code than
'the workbook it saves. A new workbook carries none of it.
'
'The four analysis sheets are named after the tags the pass-through translator
'answers with. The other five are the internal sheets the dropdown manager, the
'pivot manager, the chart registry and the spatial builder are bound to.
Private Sub NewLinelistWorkbook()
    Dim sheetNames As Variant
    Dim idx As Long

    On Error Resume Next
        If Not OutWkb Is Nothing Then DeleteWorkbook OutWkb
    On Error GoTo 0

    Set OutWkb = NewWorkbook()

    sheetNames = Array(SHEET_NORMAL, SHEET_TEMPORAL, SHEET_SPATIAL, _
                       SHEET_SPATIOTEMPORAL, SHEET_CUSTOM_CHOICE, _
                       SHEET_CUSTOM_PIVOT, SHEET_ANA_NAMES, _
                       SHEET_DROPDOWN_LISTS, SHEET_SPATIAL_TABLES)

    For idx = LBound(sheetNames) To UBound(sheetNames)
        EnsureWorksheet CStr(sheetNames(idx)), OutWkb, clearSheet:=True, _
                        visibility:=xlSheetVisible
    Next idx

    Specs.TestAssignLLWorkbook OutWkb
    Set LL = Linelist.Create(Specs)
End Sub


'@sub-title Build the time series tables of one setup table and hand back the writer.
'@param dataRows Variant. The setup rows to write and build.
'@return AnalysisOutput. The instance that ran, so its report can be read.
Private Function RunTimeSeriesBuild(ByVal dataRows As Variant) As AnalysisOutput
    Dim sut As AnalysisOutput
    Dim specSh As Worksheet

    Set specSh = BuildSetupTable(dataRows)
    NewLinelistWorkbook

    Set sut = AnalysisOutput.Create(specSh, LL)
    sut.WriteAnalysis AnalysisBuildStageTimeSeriesTables

    ' The freeze-pane pass activates worksheets of the linelist workbook, which
    ' makes that workbook the active one. The assertion harness writes its
    ' results into a worksheet of THIS workbook, and a write into a workbook
    ' that is not active raises 1004 out of PrintResults with the run already
    ' over. Handing the screen back costs one line.
    ThisWorkbook.Activate

    Set RunTimeSeriesBuild = sut
End Function

'@sub-title The report entries of one build, joined for a failure message.
'@param sut AnalysisOutput. The instance that ran.
'@return String. Every label the build filed, joined.
Private Function MilestoneLabels(ByVal sut As AnalysisOutput) As String
    Dim entries As Checking
    Dim keys As BetterArray
    Dim idx As Long
    Dim joined As String

    If sut Is Nothing Then Exit Function
    If Not sut.HasCheckings Then Exit Function

    Set entries = sut.CheckingValues
    Set keys = entries.ListOfKeys()
    If keys Is Nothing Then Exit Function

    For idx = keys.LowerBound To keys.UpperBound
        joined = joined & "[" & entries.ValueOf(CStr(keys.Item(idx)), checkingLabel) & "]"
    Next idx

    MilestoneLabels = joined
End Function

'@sub-title How many entries the navigation dropdown of one section carries.
'@param sectionId String. The table identifier of the row that opens the section.
'@return Long. The entry count, or 0 when no such dropdown was built.
Private Function SectionEntryCount(ByVal sectionId As String) As Long
    Dim entries As BetterArray

    On Error Resume Next
    Set entries = LL.Dropdown().Values(TEMPORAL_GOTO_PREFIX & "gotosection" & sectionId)
    On Error GoTo 0

    If entries Is Nothing Then Exit Function
    SectionEntryCount = entries.Length
End Function

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create rejects a Nothing specs worksheet.
'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingSpecSheet()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingSpecSheet"
    On Error GoTo TestFail

    Dim ao As AnalysisOutput

    On Error Resume Next
    Set ao = AnalysisOutput.Create(Nothing, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing specs sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects a Nothing linelist when the specs sheet is valid.
'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingLinelist()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingLinelist"
    If Not FixtureReady("TestCreateRejectsNothingLinelist") Then Exit Sub
    On Error GoTo TestFail

    Dim ao As AnalysisOutput

    On Error Resume Next
    Set ao = AnalysisOutput.Create(SpecsWkb.Worksheets(DICT_SHEET), Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing linelist should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLinelist", Err.Number, Err.Description
End Sub


'@section The section loop
'===============================================================================

'@sub-title Each section of a temporal sheet gets its own three entries.
'@details
'This is issue 337. Two sections of three tables each, and the reported symptom
'was a dropdown listing one entry fewer than the setup defines. Reading both
'dropdowns is what states that every drawn table is filed under the section it
'belongs to.
'@TestMethod("AnalysisOutput")
Public Sub TestEachSectionListsEveryTableItHolds()
    Dim sut As AnalysisOutput

    CustomTestSetTitles Assert, "AnalysisOutput", "TestEachSectionListsEveryTableItHolds"
    If Not FixtureReady("TestEachSectionListsEveryTableItHolds") Then Exit Sub
    On Error GoTo TestFail

    Set sut = RunTimeSeriesBuild(TwoWholeSections())

    Assert.AreEqual CLng(3), SectionEntryCount("TS_tab1"), _
                    "The first section holds three tables, so its dropdown " & _
                    "lists three headers"
    Assert.AreEqual CLng(3), SectionEntryCount("TS_tab4"), _
                    "And so does the second one. Listing two is what the " & _
                    "field reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEachSectionListsEveryTableItHolds", _
                         Err.Number, Err.Description
End Sub

'@sub-title Every table drawn files its own milestone entry.
'@details
'A clean analysis phase used to read the same whether it built forty tables or
'none. One entry per table drawn is what tells the two apart, and the count is
'the same number the section dropdowns list.
'@TestMethod("AnalysisOutput")
Public Sub TestEachDrawnTableFilesItsMilestone()
    Dim sut As AnalysisOutput

    CustomTestSetTitles Assert, "AnalysisOutput", "TestEachDrawnTableFilesItsMilestone"
    If Not FixtureReady("TestEachDrawnTableFilesItsMilestone") Then Exit Sub
    On Error GoTo TestFail

    Set sut = RunTimeSeriesBuild(TwoWholeSections())

    Assert.AreEqual CLng(6), sut.TablesWritten, _
                    "The two sections hold six tables between them, and each " & _
                    "drawn table files its own entry"
    Assert.IsTrue sut.HasCheckings, _
                  "A build that drew tables has something to report"
    Assert.IsTrue InStr(1, MilestoneLabels(sut), "table TS_tab1 written on ") > 0, _
                  "The entry names the table and the sheet it landed on, and the " & _
                  "report holds " & MilestoneLabels(sut)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEachDrawnTableFilesItsMilestone", _
                         Err.Number, Err.Description
End Sub

'@sub-title A table the build could not draw is left out of the count.
'@TestMethod("AnalysisOutput")
Public Sub TestARejectedTableIsLeftOutOfTheCount()
    Dim sut As AnalysisOutput

    CustomTestSetTitles Assert, "AnalysisOutput", "TestARejectedTableIsLeftOutOfTheCount"
    If Not FixtureReady("TestARejectedTableIsLeftOutOfTheCount") Then Exit Sub
    On Error GoTo TestFail

    'The same six rows, with the anchor of the second section malformed.
    Set sut = RunTimeSeriesBuild(SecondSectionWithABadAnchor())

    Assert.AreEqual CLng(5), sut.TablesWritten, _
                    "Five of the six rows were drawn, so five is the count"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARejectedTableIsLeftOutOfTheCount", _
                         Err.Number, Err.Description
End Sub

'@sub-title A section whose anchor is rejected keeps its tables to itself.
'@details
'The other half of issue 337: one section lost entries and the section above it
'gained them. The anchor of the second section names a time variable the
'dictionary does not carry, so it is never drawn, and what the two rows under it
'must not do is join the first section.
'@TestMethod("AnalysisOutput")
Public Sub TestARejectedAnchorDoesNotLendItsTablesToTheSectionAbove()
    Dim sut As AnalysisOutput

    CustomTestSetTitles Assert, "AnalysisOutput", _
                        "TestARejectedAnchorDoesNotLendItsTablesToTheSectionAbove"
    If Not FixtureReady("TestARejectedAnchorDoesNotLendItsTablesToTheSectionAbove") Then Exit Sub
    On Error GoTo TestFail

    Set sut = RunTimeSeriesBuild(SecondSectionWithABadAnchor())

    Assert.AreEqual CLng(3), SectionEntryCount("TS_tab1"), _
                    "The first section still lists its own three tables and " & _
                    "nothing else"
    Assert.AreEqual CLng(2), SectionEntryCount("TS_tab5"), _
                    "And the second section opens on the first row of it that " & _
                    "is drawn, carrying the two tables it has"
    Assert.AreEqual CLng(0), SectionEntryCount("TS_tab4"), _
                    "The rejected row anchors nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARejectedAnchorDoesNotLendItsTablesToTheSectionAbove", _
                         Err.Number, Err.Description
End Sub

'@sub-title A row the build cannot draw is skipped and the rest of the sheet is built.
'@details
'The per-table handler arms before TableSpecs.Create, so a malformed row costs
'its own table and nothing else. What matters as much is that the build carries
'on: the two rows under the rejected one are still drawn.
'@TestMethod("AnalysisOutput")
Public Sub TestARejectedRowCostsOnlyItsOwnTable()
    Dim sut As AnalysisOutput
    Dim outsh As Worksheet

    CustomTestSetTitles Assert, "AnalysisOutput", "TestARejectedRowCostsOnlyItsOwnTable"
    If Not FixtureReady("TestARejectedRowCostsOnlyItsOwnTable") Then Exit Sub
    On Error GoTo TestFail

    Set sut = RunTimeSeriesBuild(SecondSectionWithABadAnchor())
    Set outsh = OutWkb.Worksheets(SHEET_TEMPORAL)

    Assert.IsTrue RangeExistsOnSheet(outsh, "SECTION_TS_tab1"), _
                  "The first section was drawn"
    Assert.IsTrue RangeExistsOnSheet(outsh, "SECTION_TS_tab5"), _
                  "The second section was drawn from the first row of it that " & _
                  "passed validation"
    Assert.IsTrue (Not RangeExistsOnSheet(outsh, "STARTROW_TS_tab4")), _
                  "And the rejected row wrote nothing at all"
    Assert.IsTrue sut.HasCheckings, _
                  "The build filed a report"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARejectedRowCostsOnlyItsOwnTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title The section dropdown of the sheet lists one entry per section.
'@details
'The sheet-level dropdown is built from the rows that start a section, so it is
'the other reading of the same loop state.
'@TestMethod("AnalysisOutput")
Public Sub TestTheSheetDropdownListsOneEntryPerSection()
    Dim sut As AnalysisOutput
    Dim entries As BetterArray

    CustomTestSetTitles Assert, "AnalysisOutput", "TestTheSheetDropdownListsOneEntryPerSection"
    If Not FixtureReady("TestTheSheetDropdownListsOneEntryPerSection") Then Exit Sub
    On Error GoTo TestFail

    Set sut = RunTimeSeriesBuild(TwoWholeSections())

    On Error Resume Next
    Set entries = LL.Dropdown().Values(TEMPORAL_GOTO_PREFIX & "gotosection")
    On Error GoTo 0

    Assert.IsTrue (Not entries Is Nothing), _
                  "The sheet carries a section dropdown"

    If entries Is Nothing Then Exit Sub

    Assert.AreEqual CLng(2), entries.Length, _
                    "Two sections are defined, so the sheet dropdown lists two"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSheetDropdownListsOneEntryPerSection", _
                         Err.Number, Err.Description
End Sub


'@section The application state
'===============================================================================

'@sub-title WriteAnalysis hands back the state it opened on.
'@details
'The class opens an ApplicationState of its own, suppresses events for the
'length of the build and restores on the way out, so a caller reads the same
'settings after the build as before it. The known state is written on the line
'above the call, because that is the moment the snapshot is taken. It is a
'deliberately IDLE state -- screen on, events live, automatic calculation --
'since a restore can only be seen against settings the build actually changes.
'The freeze-pane pass runs inside, and each of its four activations writes the
'busy state again; those forced calls are the ones that must not leak past the
'restore.
'@TestMethod("AnalysisOutput")
Public Sub TestWriteAnalysisRestoresTheStateItOpenedOn()
    Dim sut As AnalysisOutput
    Dim specSh As Worksheet
    Dim errNumber As Long
    Dim errDesc As String

    CustomTestSetTitles Assert, "AnalysisOutput", "TestWriteAnalysisRestoresTheStateItOpenedOn"
    If Not FixtureReady("TestWriteAnalysisRestoresTheStateItOpenedOn") Then Exit Sub
    On Error GoTo TestFail

    Set specSh = BuildSetupTable(TwoWholeSections())
    NewLinelistWorkbook
    Set sut = AnalysisOutput.Create(specSh, LL)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    sut.WriteAnalysis AnalysisBuildStageTimeSeriesTables

    ' Same reason as RunTimeSeriesBuild: the freeze-pane pass leaves the
    ' linelist workbook active and the harness writes into this one.
    ThisWorkbook.Activate

    Assert.IsTrue Application.ScreenUpdating, _
                  "WriteAnalysis owns the screen for the build, so it owes it " & _
                  "back on the way out"
    Assert.IsTrue Application.EnableEvents, _
                  "And the events it suppressed"
    Assert.AreEqual CLng(xlCalculationAutomatic), CLng(Application.Calculation), _
                    "And the calculation mode it found"

    Exit Sub
TestFail:
    errNumber = Err.Number
    errDesc = Err.Description

    ' A failure inside the build can leave the screen off, and every test after
    ' this one would then run blind.
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0

    CustomTestLogFailure Assert, "TestWriteAnalysisRestoresTheStateItOpenedOn", _
                         errNumber, errDesc
End Sub


'@section Helpers used by the assertions
'===============================================================================

'@fun-title Whether a named range exists on one worksheet.
'@param sh Worksheet. The worksheet to resolve against.
'@param rngName String. The name to look up.
'@return Boolean. True when the name resolves on that worksheet.
Private Function RangeExistsOnSheet(ByVal sh As Worksheet, _
                                    ByVal rngName As String) As Boolean
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Function
    RangeExistsOnSheet = (StrComp(rng.Worksheet.Name, sh.Name, vbTextCompare) = 0)
End Function
