Attribute VB_Name = "TestAnalysisRanges"
Attribute VB_Description = "Tests for AnalysisRanges class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnalysisRanges class")

'@description
'Validates the AnalysisRanges class, which answers the named-range strings the
'analysis writers share. The class holds no worksheet and no specification, so
'this module needs no fixture beyond the output sheet.
'
'THE LITERALS ARE PINNED ON PURPOSE
'-------------------------------------------------------------------------------
'Every test below compares against a spelled-out string rather than against a
'second concatenation. A test that builds its expectation the same way the class
'does passes whatever the class says, which is no test at all. These strings are
'the contract with ranges already written onto generated workbooks, so a change
'to any one of them has to be a deliberate edit here as well.
'
'WHY THE TWO-INSTANCE TEST MATTERS
'-------------------------------------------------------------------------------
'The class carries VB_PredeclaredId, so "AnalysisRanges" names a live instance
'as well as the type. A UDT field written through the predeclared instance
'instead of a fresh one would make every table on the sheet answer with the
'last id bound. TestTwoInstancesKeepSeparateIds is what catches that.
'@depends AnalysisRanges, BetterArray, Checking, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

' Two ids that share no prefix, so a leaked field cannot pass by coincidence.
Private Const TABLE_ID As String = "BA_tab5"
Private Const SECTION_ID As String = "TS_tab2"

Private Assert As CustomTest


'@section Lifecycle
'===============================================================================

'@sub-title Create the assert object for the module.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisRanges"
End Sub

'@sub-title Print results and tear down module-level fixtures.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updating before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Factory
'===============================================================================

'@sub-title Verify Create refuses an empty identifier.
'@details
'An empty id builds "VALUES_COL_1_", a legal name every table would claim in
'turn, so the sheet would end up with names pointing at the wrong table and no
'error anywhere.
'@TestMethod("AnalysisRanges")
Public Sub TestCreateRejectsAnEmptyId()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestCreateRejectsAnEmptyId"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Dim errNumber As Long

    On Error Resume Next
    Set rangeNames = AnalysisRanges.Create(vbNullString)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "An empty range id should raise InvalidArgument"
    Assert.IsTrue (rangeNames Is Nothing), _
                  "Create with an empty id should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsAnEmptyId", Err.Number, Err.Description
End Sub

'@sub-title Verify Id answers the identifier it was given.
'@TestMethod("AnalysisRanges")
Public Sub TestIdAnswersTheIdentifierItWasGiven()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestIdAnswersTheIdentifierItWasGiven"
    On Error GoTo TestFail

    Assert.AreEqual TABLE_ID, AnalysisRanges.Create(TABLE_ID).Id, _
                    "Id should answer the string passed to Create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIdAnswersTheIdentifierItWasGiven", Err.Number, Err.Description
End Sub

'@sub-title Verify two instances keep separate identifiers.
'@details
'The class is predeclared, so a field written on the wrong instance would make
'every table answer with the id bound last. Both instances are held at once and
'read after the second is built, which is the only ordering that catches it.
'@TestMethod("AnalysisRanges")
Public Sub TestTwoInstancesKeepSeparateIds()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTwoInstancesKeepSeparateIds"
    On Error GoTo TestFail

    Dim tabNames As AnalysisRanges
    Dim secNames As AnalysisRanges

    Set tabNames = AnalysisRanges.Create(TABLE_ID)
    Set secNames = AnalysisRanges.Create(SECTION_ID)

    Assert.AreEqual TABLE_ID, tabNames.Id, _
                    "The first instance should keep its own id"
    Assert.AreEqual SECTION_ID, secNames.Id, _
                    "The second instance should keep its own id"
    Assert.AreEqual "ROW_CATEGORIES_BA_tab5", tabNames.RowCategories, _
                    "The first instance should still build names from its own id"
    Assert.AreEqual "ROW_CATEGORIES_TS_tab2", secNames.RowCategories, _
                    "The second instance should build names from its own id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoInstancesKeepSeparateIds", Err.Number, Err.Description
End Sub


'@section Column families
'===============================================================================

'@sub-title Verify the three column families number the column and carry the id.
'@TestMethod("AnalysisRanges")
Public Sub TestColumnFamiliesNumberTheColumn()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestColumnFamiliesNumberTheColumn"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(TABLE_ID)

    Assert.AreEqual "VALUES_COL_1_BA_tab5", rangeNames.ValuesCol(1), _
                    "ValuesCol should spell the value column of column one"
    Assert.AreEqual "VALUES_COL_12_BA_tab5", rangeNames.ValuesCol(12), _
                    "ValuesCol should carry a two-digit column number unpadded"
    Assert.AreEqual "LABEL_COL_1_BA_tab5", rangeNames.LabelCol(1), _
                    "LabelCol should spell the label column of column one"
    Assert.AreEqual "PERC_COL_1_BA_tab5", rangeNames.PercCol(1), _
                    "PercCol should spell the percentage column of column one"
    Assert.AreEqual "PERC_COL_3_BA_tab5", rangeNames.PercCol(3), _
                    "PercCol should spell the percentage column of column three"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestColumnFamiliesNumberTheColumn", Err.Number, Err.Description
End Sub


'@section Spatial input naming
'===============================================================================

'@sub-title Verify the geographic spatial type gives the geographic tag.
'@TestMethod("AnalysisRanges")
Public Sub TestSpatialInputUsesTheGeographicTag()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestSpatialInputUsesTheGeographicTag"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create("SPT_tab1")

    Assert.AreEqual "INPUTSPTGEO_", rangeNames.SpatialInputTag("geo"), _
                    "A geo table should use the geographic tag"
    Assert.AreEqual "INPUTSPTGEO_2_SPT_tab1", rangeNames.SpatialInput(2, "geo"), _
                    "A geo input cell should be numbered and carry the id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialInputUsesTheGeographicTag", Err.Number, Err.Description
End Sub

'@sub-title Verify the health facility spatial type gives the facility tag.
'@details
'This is the half of the pair that a table writer reading an unpopulated setup
'column never reached, so the facility cells were named with the geographic tag
'while the label formulas referenced the facility one.
'@TestMethod("AnalysisRanges")
Public Sub TestSpatialInputUsesTheFacilityTag()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestSpatialInputUsesTheFacilityTag"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create("SPT_tab1")

    Assert.AreEqual "INPUTSPTHF_", rangeNames.SpatialInputTag("hf"), _
                    "A facility table should use the facility tag"
    Assert.AreEqual "INPUTSPTHF_2_SPT_tab1", rangeNames.SpatialInput(2, "hf"), _
                    "A facility input cell should be numbered and carry the id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialInputUsesTheFacilityTag", Err.Number, Err.Description
End Sub

'@sub-title Verify the spatial type is matched without regard to case.
'@details
'The value reaches this class from a dictionary probe today, but it began life
'as a value a user typed into a setup column, and the two spellings should not
'name cells differently.
'@TestMethod("AnalysisRanges")
Public Sub TestSpatialInputIgnoresTheCaseOfTheType()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestSpatialInputIgnoresTheCaseOfTheType"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create("SPT_tab1")

    Assert.AreEqual "INPUTSPTHF_", rangeNames.SpatialInputTag("HF"), _
                    "An upper case facility type should give the facility tag"
    Assert.AreEqual "INPUTSPTGEO_", rangeNames.SpatialInputTag("Geo"), _
                    "A mixed case geo type should give the geographic tag"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialInputIgnoresTheCaseOfTheType", Err.Number, Err.Description
End Sub

'@sub-title Verify an unrecognised spatial type raises instead of defaulting.
'@details
'Answering the geographic tag for an unknown type is exactly the behaviour that
'hid the disagreement: a table whose type could not be read was named as though
'it were geographic, and nothing said so. An empty string is the case that
'actually occurred.
'@TestMethod("AnalysisRanges")
Public Sub TestSpatialInputRejectsAnUnknownType()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestSpatialInputRejectsAnUnknownType"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Dim emptyErr As Long
    Dim otherErr As Long
    Dim builtName As String

    Set rangeNames = AnalysisRanges.Create("SPT_tab1")

    On Error Resume Next
    builtName = rangeNames.SpatialInputTag(vbNullString)
    emptyErr = Err.Number
    Err.Clear
    builtName = rangeNames.SpatialInput(1, "adm1")
    otherErr = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, emptyErr, _
                    "An empty spatial type should raise rather than assume geographic"
    Assert.AreEqual ProjectError.InvalidArgument, otherErr, _
                    "An unrecognised spatial type should raise"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialInputRejectsAnUnknownType", Err.Number, Err.Description
End Sub


'@section Category, total and missing families
'===============================================================================

'@sub-title Verify the category and start-row names.
'@TestMethod("AnalysisRanges")
Public Sub TestCategoryAndStartRowNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestCategoryAndStartRowNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(TABLE_ID)

    Assert.AreEqual "ROW_CATEGORIES_BA_tab5", rangeNames.RowCategories, _
                    "RowCategories should spell the row category band"
    Assert.AreEqual "STARTROW_BA_tab5", rangeNames.StartRow, _
                    "StartRow should spell the start row marker"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoryAndStartRowNames", Err.Number, Err.Description
End Sub

'@sub-title Verify the total and missing band names.
'@TestMethod("AnalysisRanges")
Public Sub TestTotalAndMissingBandNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTotalAndMissingBandNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(TABLE_ID)

    Assert.AreEqual "TOTAL_ROW_BA_tab5", rangeNames.TotalRow, _
                    "TotalRow should spell the total row band"
    Assert.AreEqual "TOTAL_ROW_VALUES_BA_tab5", rangeNames.TotalRowValues, _
                    "TotalRowValues should spell the values inside the total row"
    Assert.AreEqual "TOTAL_COL_VALUES_BA_tab5", rangeNames.TotalColValues, _
                    "TotalColValues should spell the values inside the total column"
    Assert.AreEqual "MISSING_ROW_VALUES_BA_tab5", rangeNames.MissingRowValues, _
                    "MissingRowValues should spell the values inside the missing row"
    Assert.AreEqual "MISSING_COL_VALUES_BA_tab5", rangeNames.MissingColValues, _
                    "MissingColValues should spell the values inside the missing column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalAndMissingBandNames", Err.Number, Err.Description
End Sub

'@sub-title Verify the four names where the total and missing bands cross.
'@details
'These four are the corner cells, and every one of them had two builders: the
'table writer names them and the formula writer fills them.
'@TestMethod("AnalysisRanges")
Public Sub TestTotalAndMissingIntersectionNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTotalAndMissingIntersectionNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(TABLE_ID)

    Assert.AreEqual "TOTAL_TOTAL_BA_tab5", rangeNames.TotalTotal, _
                    "TotalTotal should spell the total row and total column corner"
    Assert.AreEqual "TOTAL_MISSING_BA_tab5", rangeNames.TotalMissing, _
                    "TotalMissing should spell the total row and missing column corner"
    Assert.AreEqual "MISSING_TOTAL_BA_tab5", rangeNames.MissingTotal, _
                    "MissingTotal should spell the missing row and total column corner"
    Assert.AreEqual "MISSING_MISSING_BA_tab5", rangeNames.MissingMissing, _
                    "MissingMissing should spell the missing row and missing column corner"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalAndMissingIntersectionNames", Err.Number, Err.Description
End Sub


'@section Temporal and spatial control families
'===============================================================================

'@sub-title Verify the date and period names.
'@TestMethod("AnalysisRanges")
Public Sub TestTemporalNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTemporalNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(SECTION_ID)

    Assert.AreEqual "TIME_UNIT_TS_tab2", rangeNames.TimeUnit, _
                    "TimeUnit should spell the time unit control"
    Assert.AreEqual "START_DATE_TS_tab2", rangeNames.StartDate, _
                    "StartDate should spell the start date cell"
    Assert.AreEqual "END_DATE_TS_tab2", rangeNames.EndDate, _
                    "EndDate should spell the end date cell"
    Assert.AreEqual "START_TIME_PERIOD_TS_tab2", rangeNames.StartTimePeriod, _
                    "StartTimePeriod should spell the period start marker"
    Assert.AreEqual "END_TIME_PERIOD_TS_tab2", rangeNames.EndTimePeriod, _
                    "EndTimePeriod should spell the period end marker"
    Assert.AreEqual "FIRST_VALUE_START_TIME_TS_tab2", rangeNames.FirstValueStartTime, _
                    "FirstValueStartTime should spell the first value of the period"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTemporalNames", Err.Number, Err.Description
End Sub

'@sub-title Verify the period information and date validation names.
'@TestMethod("AnalysisRanges")
Public Sub TestPeriodInformationAndValidationNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestPeriodInformationAndValidationNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(SECTION_ID)

    Assert.AreEqual "INFO_START_DATE_TS_tab2", rangeNames.InfoStartDate, _
                    "InfoStartDate should spell the start date information cell"
    Assert.AreEqual "INFO_END_DATE_TS_tab2", rangeNames.InfoEndDate, _
                    "InfoEndDate should spell the end date information cell"
    Assert.AreEqual "INFO_ANA_PERIOD_TS_tab2", rangeNames.InfoAnaPeriod, _
                    "InfoAnaPeriod should spell the analysis period information cell"
    Assert.AreEqual "VALIDATION_MIN_DATE_TS_tab2", rangeNames.ValidationMinDate, _
                    "ValidationMinDate should spell the minimum date validation cell"
    Assert.AreEqual "VALIDATION_MAX_DATE_TS_tab2", rangeNames.ValidationMaxDate, _
                    "ValidationMaxDate should spell the maximum date validation cell"

    ' The two cells the reader types into. Six formulas of the formula writer
    ' read them and the table writer names them, so they are a shared family.
    Assert.AreEqual "USER_START_DATE_TS_tab2", rangeNames.UserStartDate, _
                    "UserStartDate should spell the cell the reader types the first date into"
    Assert.AreEqual "USER_END_DATE_TS_tab2", rangeNames.UserEndDate, _
                    "UserEndDate should spell the cell the reader types the last date into"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPeriodInformationAndValidationNames", Err.Number, Err.Description
End Sub

'@sub-title Verify the two time unit lists and the scope that picks between them.
'@details
'The spatio-temporal tables live on a worksheet of their own and read a time unit
'list of their own. The table writer built that name and the formula writer asked
'for the plain one, so no spatio-temporal table ever got its dropdown. One member
'answers for both callers now.
'@TestMethod("AnalysisRanges")
Public Sub TestTheTwoTimeUnitListsAndTheScopeThatPicks()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTheTwoTimeUnitListsAndTheScopeThatPicks"
    On Error GoTo TestFail

    Assert.AreEqual "TIME_UNIT_LIST", AnalysisRanges.TimeUnitList, _
                    "TimeUnitList should spell the shared time unit list"
    Assert.AreEqual "SPTIME_UNIT_LIST", AnalysisRanges.SpatioTemporalTimeUnitList, _
                    "The spatio-temporal list carries the SP prefix"
    Assert.AreEqual "TIME_UNIT_LIST", AnalysisRanges.TimeUnitListOfScope(False), _
                    "A time series table reads the shared list"
    Assert.AreEqual "SPTIME_UNIT_LIST", AnalysisRanges.TimeUnitListOfScope(True), _
                    "A spatio-temporal table reads its own list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTwoTimeUnitListsAndTheScopeThatPicks", Err.Number, Err.Description
End Sub

'@sub-title Verify the administrative dropdown and population divisor names.
'@TestMethod("AnalysisRanges")
Public Sub TestSpatialControlNames()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestSpatialControlNames"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create("SPT_tab1")

    Assert.AreEqual "ADM_DROPDOWN_SPT_tab1", rangeNames.AdmDropdown, _
                    "AdmDropdown should spell the administrative level dropdown"
    Assert.AreEqual "DEVIDEPOP_SPT_tab1", rangeNames.DevidePop, _
                    "DevidePop should spell the population divisor control"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialControlNames", Err.Number, Err.Description
End Sub


'@section Workbook-wide lists
'===============================================================================

'@sub-title Verify the workbook-wide lists carry no identifier.
'@details
'One copy of each of these serves every table, so appending an id would give
'each table a list of its own and break the sharing. They are read off the
'predeclared instance without calling Create, which is the only way a caller
'holding no id can reach them.
'@TestMethod("AnalysisRanges")
Public Sub TestWorkbookWideListsCarryNoIdentifier()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestWorkbookWideListsCarryNoIdentifier"
    On Error GoTo TestFail

    Assert.AreEqual "TIME_UNIT_LIST", AnalysisRanges.TimeUnitList, _
                    "TimeUnitList should carry no identifier"
    Assert.AreEqual "ADM_UNIT_LIST", AnalysisRanges.AdmUnitList, _
                    "AdmUnitList should carry no identifier"
    Assert.AreEqual "POPULATION_FACTOR_LIST", AnalysisRanges.PopulationFactorList, _
                    "PopulationFactorList should carry no identifier"

    ' A bound instance answers the same strings, so a caller that happens to
    ' hold one does not get a second, table-scoped list by accident.
    Assert.AreEqual "TIME_UNIT_LIST", AnalysisRanges.Create(TABLE_ID).TimeUnitList, _
                    "A bound instance should answer the same shared list name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWorkbookWideListsCarryNoIdentifier", Err.Number, Err.Description
End Sub

'@sub-title Verify the time unit list name differs from the per-table time unit name.
'@details
'"TIME_UNIT_LIST" and "TIME_UNIT_" & id are two different ranges, and the first
'is a prefix of nothing the second builds, so a partial match cannot confuse
'them. Stated as a test because the two names read alike.
'@TestMethod("AnalysisRanges")
Public Sub TestTheSharedListIsNotThePerTableTimeUnit()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestTheSharedListIsNotThePerTableTimeUnit"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Set rangeNames = AnalysisRanges.Create(SECTION_ID)

    Assert.IsTrue (rangeNames.TimeUnit <> rangeNames.TimeUnitList), _
                  "The per-table time unit and the shared list are different ranges"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSharedListIsNotThePerTableTimeUnit", Err.Number, Err.Description
End Sub


'@section Whole-surface properties
'===============================================================================

'@sub-title Verify every identifier-suffixed name ends with the identifier.
'@details
'A sweep over the whole surface rather than one assertion per member. A new
'member that forgets the id, or spells the separator differently, is caught here
'without the test having to be extended.
'@TestMethod("AnalysisRanges")
Public Sub TestEveryIdSuffixedNameEndsWithTheId()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestEveryIdSuffixedNameEndsWithTheId"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Dim built As BetterArray
    Dim counter As Long
    Dim builtName As String
    Dim tailLength As Long

    Set rangeNames = AnalysisRanges.Create(TABLE_ID)
    tailLength = Len("_" & TABLE_ID)

    Set built = New BetterArray
    built.LowerBound = 1
    PushEveryIdSuffixedName built, rangeNames

    For counter = built.LowerBound To built.UpperBound
        builtName = CStr(built.Item(counter))
        Assert.AreEqual "_" & TABLE_ID, Right$(builtName, tailLength), _
                        builtName & " should end with an underscore and the id"
    Next counter

    Assert.LogSuccesses "AnalysisRanges built " & built.Length & _
                        " id-suffixed names for " & TABLE_ID

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEveryIdSuffixedNameEndsWithTheId", Err.Number, Err.Description
End Sub

'@sub-title Verify no two members answer the same string.
'@details
'Two members answering one name would have two ranges fighting over one cell,
'and the second Name assignment would silently move the first. A copy and paste
'slip inside this class is the likely cause, so the check is a sweep.
'@TestMethod("AnalysisRanges")
Public Sub TestNoTwoNamesCollide()
    CustomTestSetTitles Assert, "AnalysisRanges", "TestNoTwoNamesCollide"
    On Error GoTo TestFail

    Dim rangeNames As AnalysisRanges
    Dim built As BetterArray
    Dim outer As Long
    Dim inner As Long
    Dim collisions As Long

    Set rangeNames = AnalysisRanges.Create(TABLE_ID)

    Set built = New BetterArray
    built.LowerBound = 1
    PushEveryIdSuffixedName built, rangeNames
    built.Push rangeNames.TimeUnitList, rangeNames.AdmUnitList, _
               rangeNames.PopulationFactorList

    collisions = 0
    For outer = built.LowerBound To built.UpperBound - 1
        For inner = outer + 1 To built.UpperBound
            If CStr(built.Item(outer)) = CStr(built.Item(inner)) Then
                collisions = collisions + 1
            End If
        Next inner
    Next outer

    Assert.AreEqual CLng(0), collisions, _
                    "No two members of AnalysisRanges should answer the same name"
    Assert.LogSuccesses "Checked " & built.Length & " names for collisions"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNoTwoNamesCollide", Err.Number, Err.Description
End Sub


'@section Shared helpers
'===============================================================================

'@sub-title Push every identifier-suffixed name onto an array.
'@details
'The two sweeps above both need the whole surface, and a member added to the
'class but to only one of the two lists would leave one sweep quietly weaker
'than it reads. One list serves both, so a new member is added in one place.
'@param built BetterArray. The array to push onto.
'@param rangeNames AnalysisRanges. The bound instance to read.
Private Sub PushEveryIdSuffixedName(ByVal built As BetterArray, _
                                    ByVal rangeNames As AnalysisRanges)
    built.Push rangeNames.ValuesCol(1), rangeNames.LabelCol(1), _
               rangeNames.PercCol(1), rangeNames.PercLabelCol, _
               rangeNames.SpatialInput(1, "geo"), rangeNames.SpatialInput(1, "hf"), _
               rangeNames.RowCategories, rangeNames.LabelRowCategories, _
               rangeNames.ColumnCategories, rangeNames.InteriorValues, _
               rangeNames.OuterValues, rangeNames.Section, rangeNames.EndTable, _
               rangeNames.StartRow, rangeNames.StartCol

    built.Push rangeNames.TotalRow, rangeNames.TotalCol, _
               rangeNames.TotalRowValues, rangeNames.TotalColValues, _
               rangeNames.MissingRow, rangeNames.MissingCol, _
               rangeNames.MissingRowValues, rangeNames.MissingColValues, _
               rangeNames.TotalTotal, rangeNames.TotalMissing, _
               rangeNames.MissingTotal, rangeNames.MissingMissing

    built.Push rangeNames.TimeUnit, rangeNames.StartDate, rangeNames.EndDate, _
               rangeNames.StartTimePeriod, rangeNames.EndTimePeriod, _
               rangeNames.FirstValueStartTime, rangeNames.InfoStartDate, _
               rangeNames.InfoEndDate, rangeNames.InfoAnaPeriod, _
               rangeNames.ValidationMinDate, rangeNames.ValidationMaxDate, _
               rangeNames.MinMinDate, rangeNames.MaxMaxDate, _
               rangeNames.UserStartDate, rangeNames.UserEndDate

    built.Push rangeNames.AdmDropdown, rangeNames.DevidePop, _
               rangeNames.PopFact, rangeNames.PopFactLabel, _
               rangeNames.PopPrevFact, rangeNames.PreviousAdm
End Sub
