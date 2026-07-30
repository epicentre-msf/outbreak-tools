Attribute VB_Name = "TestTemporalSection"
Attribute VB_Description = "Tests for TemporalSection class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for TemporalSection class")

'@description
'Drives TemporalSection, the holder AnalysisOutput carries while it writes the
'tables of a time series or spatio-temporal sheet. Five defects came out of the
'look-ahead this class replaced, and five of these tests are those defects
'written down: a two-section run writes both pairs of date bounds, a section
'lists only its own headers, the last table of the last section is inside the
'bounds, a section holding one table pushes its header once, and a run with no
'table writes nothing and raises nothing.
'
'The class holds strings and a list, so the whole suite runs with no workbook,
'no linelist and no dictionary.
'@depends TemporalSection, CustomTest, BetterArray

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const TABLE_ONE As String = "ftable1"
Private Const DATE_COLUMN As String = "date_v1"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTemporalSection"
End Sub

'@sub-title Print results and tear down shared state.
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

'@section Fixture helpers
'===============================================================================

'@sub-title Add one table to the open section, named after its index.
'@param section TemporalSection. The holder under test.
'@param tabId String. Identifier of the table.
Private Sub AddOneTable(ByVal section As TemporalSection, ByVal tabId As String)
    section.AddTable tabId:=tabId, tableName:=TABLE_ONE, rowVar:=DATE_COLUMN, _
                     header:="Go to header: " & tabId
End Sub

'@sub-title How many entries of a list hold a given text.
'@param entries BetterArray. The list to read.
'@param needle String. The text to count.
'@return Long. The number of entries holding the text.
Private Function CountOf(ByVal entries As BetterArray, ByVal needle As String) As Long
    Dim idx As Long
    Dim found As Long

    If entries.Length = 0 Then Exit Function

    For idx = entries.LowerBound To entries.UpperBound
        If InStr(1, CStr(entries.Item(idx)), needle, vbTextCompare) > 0 Then
            found = found + 1
        End If
    Next idx

    CountOf = found
End Function

'@section A section that holds nothing
'===============================================================================

'@sub-title Verify a new holder has no section open and nothing to write.
'@TestMethod("TemporalSection")
Public Sub TestANewHolderIsClosedAndEmpty()
    CustomTestSetTitles Assert, "TemporalSection", "TestANewHolderIsClosedAndEmpty"
    On Error GoTo TestFail

    Dim section As TemporalSection

    Set section = TemporalSection.Create()

    Assert.IsTrue (Not section.IsOpen()), "A new holder has no section open"
    Assert.AreEqual CLng(0), section.TableCount(), "A new holder counts no table"
    Assert.AreEqual vbNullString, section.MinFormula(), _
                    "A holder with no table has no earliest date to write"
    Assert.AreEqual vbNullString, section.MaxFormula(), _
                    "A holder with no table has no latest date to write"
    Assert.AreEqual CLng(0), section.Headers().Length, _
                    "A holder with no table offers no dropdown entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestANewHolderIsClosedAndEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verify a section with no table gives no formula to write.
'@details
'This is what a run over a ListObject whose rows are all invalid leaves behind.
'The caller writes nothing and nothing raises.
'@TestMethod("TemporalSection")
Public Sub TestAnEmptySectionWritesNothing()
    CustomTestSetTitles Assert, "TemporalSection", "TestAnEmptySectionWritesNothing"
    On Error GoTo TestFail

    Dim section As TemporalSection

    Set section = TemporalSection.Create()
    section.StartSection "SEC1"
    section.EndSection

    Assert.AreEqual "SEC1", section.SectionId(), "The section keeps its identifier"
    Assert.AreEqual vbNullString, section.MinFormula(), _
                    "A section with no table has no earliest date"
    Assert.AreEqual vbNullString, section.MaxFormula(), _
                    "A section with no table has no latest date"
    Assert.AreEqual vbNullString, section.LastTableId(), _
                    "A section with no table ends on no table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnEmptySectionWritesNothing", Err.Number, Err.Description
End Sub

'@section The date bounds
'===============================================================================

'@sub-title Verify a section holding one table wraps that one table.
'@TestMethod("TemporalSection")
Public Sub TestOneTableGivesOneBound()
    CustomTestSetTitles Assert, "TemporalSection", "TestOneTableGivesOneBound"
    On Error GoTo TestFail

    Dim section As TemporalSection

    Set section = TemporalSection.Create()
    section.StartSection "T1"
    AddOneTable section, "T1"

    Assert.AreEqual "= MIN(MIN(" & TABLE_ONE & "[" & DATE_COLUMN & "]))", _
                    section.MinFormula(), "One table gives one MIN inside the wrap"
    Assert.AreEqual "= MAX(MAX(" & TABLE_ONE & "[" & DATE_COLUMN & "]))", _
                    section.MaxFormula(), "One table gives one MAX inside the wrap"
    Assert.AreEqual CLng(1), section.TableCount(), "One table is counted once"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOneTableGivesOneBound", Err.Number, Err.Description
End Sub

'@sub-title Verify the last table of a section is inside its date bounds.
'@details
'The look-ahead this class replaced stopped collecting one table short: the
'iteration that wrote the formula was the one whose own dates were left out.
'@TestMethod("TemporalSection")
Public Sub TestTheLastTableIsInsideTheBounds()
    CustomTestSetTitles Assert, "TemporalSection", "TestTheLastTableIsInsideTheBounds"
    On Error GoTo TestFail

    Dim section As TemporalSection

    Set section = TemporalSection.Create()
    section.StartSection "A1"
    section.AddTable "A1", "ftableA", "date_a", "Go to header: A1"
    section.AddTable "A2", "ftableB", "date_b", "Go to header: A2"
    section.AddTable "A3", "ftableC", "date_c", "Go to header: A3"
    section.EndSection

    Assert.AreEqual CLng(3), section.TableCount(), "Three tables are counted three times"
    Assert.AreEqual "= MIN(MIN(ftableA[date_a]), MIN(ftableB[date_b]), MIN(ftableC[date_c]))", _
                    section.MinFormula(), "Every table of the section is inside the MIN"
    Assert.AreEqual "= MAX(MAX(ftableA[date_a]), MAX(ftableB[date_b]), MAX(ftableC[date_c]))", _
                    section.MaxFormula(), "Every table of the section is inside the MAX"
    Assert.AreEqual "A3", section.LastTableId(), "The section ends on its last table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheLastTableIsInsideTheBounds", Err.Number, Err.Description
End Sub

'@sub-title Verify two sections give two separate pairs of date bounds.
'@details
'The branch that wrote the bounds of the section before the current one tested
'a condition that could never be True, so a run over several sections wrote one
'pair for the last section and nothing for the others.
'@TestMethod("TemporalSection")
Public Sub TestTwoSectionsGiveTwoPairsOfBounds()
    CustomTestSetTitles Assert, "TemporalSection", "TestTwoSectionsGiveTwoPairsOfBounds"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim firstMin As String
    Dim firstMax As String
    Dim firstId As String

    Set section = TemporalSection.Create()

    section.StartSection "A1"
    section.AddTable "A1", "ftableA", "date_a", "Go to header: A1"
    section.AddTable "A2", "ftableB", "date_b", "Go to header: A2"
    section.EndSection

    firstId = section.SectionId()
    firstMin = section.MinFormula()
    firstMax = section.MaxFormula()

    section.StartSection "B1"
    section.AddTable "B1", "ftableC", "date_c", "Go to header: B1"
    section.EndSection

    Assert.AreEqual "A1", firstId, "The first section is named after its first table"
    Assert.AreEqual "= MIN(MIN(ftableA[date_a]), MIN(ftableB[date_b]))", firstMin, _
                    "The first section carries both of its tables"
    Assert.AreEqual "= MAX(MAX(ftableA[date_a]), MAX(ftableB[date_b]))", firstMax, _
                    "The first section carries both of its tables"

    Assert.AreEqual "B1", section.SectionId(), "The second section is named after its own table"
    Assert.AreEqual "= MIN(MIN(ftableC[date_c]))", section.MinFormula(), _
                    "The second section carries its own table alone"
    Assert.AreEqual "= MAX(MAX(ftableC[date_c]))", section.MaxFormula(), _
                    "The second section carries its own table alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoSectionsGiveTwoPairsOfBounds", Err.Number, Err.Description
End Sub

'@section The dropdown entries
'===============================================================================

'@sub-title Verify a section lists only the headers of its own tables.
'@details
'The header list was cleared inside a branch that never ran, so the one
'dropdown that did get built listed every header of the sheet.
'@TestMethod("TemporalSection")
Public Sub TestASectionListsOnlyItsOwnHeaders()
    CustomTestSetTitles Assert, "TemporalSection", "TestASectionListsOnlyItsOwnHeaders"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim entries As BetterArray

    Set section = TemporalSection.Create()

    section.StartSection "A1"
    AddOneTable section, "A1"
    AddOneTable section, "A2"
    AddOneTable section, "A3"
    section.EndSection

    section.StartSection "B1"
    AddOneTable section, "B1"
    AddOneTable section, "B2"
    section.EndSection

    Set entries = section.Headers()

    Assert.AreEqual CLng(2), entries.Length, "The second section offers its two headers"
    Assert.AreEqual CLng(1), CountOf(entries, "B1"), "B1 is offered once"
    Assert.AreEqual CLng(1), CountOf(entries, "B2"), "B2 is offered once"
    Assert.AreEqual CLng(0), CountOf(entries, "A1"), _
                    "No header of the section before it survives"
    Assert.AreEqual CLng(0), CountOf(entries, "A3"), _
                    "No header of the section before it survives"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASectionListsOnlyItsOwnHeaders", Err.Number, Err.Description
End Sub

'@sub-title Verify a section holding one table offers its header once.
'@details
'A section that both started and ended on the same row pushed its header
'through two branches, and the dropdown showed the entry twice.
'@TestMethod("TemporalSection")
Public Sub TestASingleTableSectionOffersOneHeader()
    CustomTestSetTitles Assert, "TemporalSection", "TestASingleTableSectionOffersOneHeader"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim entries As BetterArray

    Set section = TemporalSection.Create()
    section.StartSection "C1"
    AddOneTable section, "C1"
    section.EndSection

    Set entries = section.Headers()

    Assert.AreEqual CLng(1), entries.Length, "One table gives one dropdown entry"
    Assert.AreEqual CLng(1), CountOf(entries, "C1"), "The one entry is offered once"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASingleTableSectionOffersOneHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify the header list handed out is a copy.
'@details
'The caller passes it straight to the dropdown builder, so growing it must not
'grow the section's own list.
'@TestMethod("TemporalSection")
Public Sub TestTheHeaderListHandedOutIsACopy()
    CustomTestSetTitles Assert, "TemporalSection", "TestTheHeaderListHandedOutIsACopy"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim entries As BetterArray

    Set section = TemporalSection.Create()
    section.StartSection "D1"
    AddOneTable section, "D1"

    Set entries = section.Headers()
    entries.Push "an entry the caller added"

    Assert.AreEqual CLng(2), entries.Length, "The caller grew the copy it holds"
    Assert.AreEqual CLng(1), section.Headers().Length, _
                    "The section still offers its one entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHeaderListHandedOutIsACopy", Err.Number, Err.Description
End Sub

'@section The state the class refuses
'===============================================================================

'@sub-title Verify a second start without an end is refused.
'@TestMethod("TemporalSection")
Public Sub TestStartingAnOpenSectionIsRefused()
    CustomTestSetTitles Assert, "TemporalSection", "TestStartingAnOpenSectionIsRefused"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim errNumber As Long

    Set section = TemporalSection.Create()
    section.StartSection "E1"

    On Error Resume Next
    section.StartSection "E2"
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "Starting a section over an open one should raise"
    Assert.AreEqual "E1", section.SectionId(), "The open section is untouched"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestStartingAnOpenSectionIsRefused", Err.Number, Err.Description
End Sub

'@sub-title Verify a section with no identifier is refused.
'@TestMethod("TemporalSection")
Public Sub TestASectionNeedsAnIdentifier()
    CustomTestSetTitles Assert, "TemporalSection", "TestASectionNeedsAnIdentifier"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim errNumber As Long

    Set section = TemporalSection.Create()

    On Error Resume Next
    section.StartSection vbNullString
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A section with no identifier should raise"
    Assert.IsTrue (Not section.IsOpen()), "Nothing was opened"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASectionNeedsAnIdentifier", Err.Number, Err.Description
End Sub

'@sub-title Verify a table added outside a section is refused.
'@TestMethod("TemporalSection")
Public Sub TestATableNeedsAnOpenSection()
    CustomTestSetTitles Assert, "TemporalSection", "TestATableNeedsAnOpenSection"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim errNumber As Long

    Set section = TemporalSection.Create()

    On Error Resume Next
    AddOneTable section, "F1"
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "A table with no section to belong to should raise"
    Assert.AreEqual CLng(0), section.TableCount(), "Nothing was counted"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATableNeedsAnOpenSection", Err.Number, Err.Description
End Sub

'@sub-title Verify ending a section that was never started is refused.
'@TestMethod("TemporalSection")
Public Sub TestEndingAClosedSectionIsRefused()
    CustomTestSetTitles Assert, "TemporalSection", "TestEndingAClosedSectionIsRefused"
    On Error GoTo TestFail

    Dim section As TemporalSection
    Dim errNumber As Long

    Set section = TemporalSection.Create()

    On Error Resume Next
    section.EndSection
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "Ending a section that was never started should raise"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEndingAClosedSectionIsRefused", Err.Number, Err.Description
End Sub

'@section The values a closed section keeps
'===============================================================================

'@sub-title Verify a closed section still answers everything the caller writes.
'@details
'The caller ends the section and then writes what it held, so ending it must
'not clear it. Starting the next one is what clears it.
'@TestMethod("TemporalSection")
Public Sub TestAClosedSectionKeepsItsValues()
    CustomTestSetTitles Assert, "TemporalSection", "TestAClosedSectionKeepsItsValues"
    On Error GoTo TestFail

    Dim section As TemporalSection

    Set section = TemporalSection.Create()
    section.StartSection "G1"
    AddOneTable section, "G1"
    AddOneTable section, "G2"
    section.EndSection

    Assert.IsTrue (Not section.IsOpen()), "The section is closed"
    Assert.AreEqual "G1", section.SectionId(), "It still answers its identifier"
    Assert.AreEqual "G2", section.LastTableId(), "It still answers its last table"
    Assert.AreEqual CLng(2), section.TableCount(), "It still counts its tables"
    Assert.AreEqual CLng(2), section.Headers().Length, "It still offers its entries"
    Assert.IsTrue (LenB(section.MinFormula()) > 0), "It still answers its earliest date"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAClosedSectionKeepsItsValues", Err.Number, Err.Description
End Sub
