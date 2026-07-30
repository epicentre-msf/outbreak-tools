Attribute VB_Name = "TestAnaTabIds"
Attribute VB_Description = "Tests for AnaTabIds class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnaTabIds class")

'@description
'Validates AnaTabIds, which records the charts of the analysis sheets and
'copies the named ranges those charts read onto the sheets the export creates.
'
'THREE SHEETS MAKE THE FIXTURE
'-------------------------------------------------------------------------------
'One sheet carries the registry table, one stands in for an analysis sheet with
'its named ranges, and one stands in for the sheet the export writes. The output
'sheet is visible because WriteGraphs calls Application.GoTo, which needs a
'selectable sheet.
'
'THE SHEETS ARE BUILT ONCE AND RESET IN A BOUNDED BLOCK
'-------------------------------------------------------------------------------
'Clearing a whole worksheet costs seconds, and a suite that did it per test took
'a green run past the runner cap. The three sheets are created once and each test
'resets a block bigger than anything the tests write.
'
'A NAME ON THE INPUT SHEET IS WORKSHEET-SCOPED
'-------------------------------------------------------------------------------
'The tests run inside one workbook, so recreating a workbook-scoped name on the
'output sheet moves the one definition there. The fixture names are
'worksheet-scoped, which also exercises the qualifier stripping: the names
'collection reports them as "sheet!NAME".
'@depends AnaTabIds, Graphs, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const REGISTRY_SHEET As String = "AnaIdsRegistry"
Private Const INPUT_SHEET As String = "AnaIdsInput"
Private Const OUTPUT_SHEET As String = "AnaIdsOutput"
Private Const REGISTRY_TABLE As String = "ana_registry"
Private Const FIXTURE_BLOCK As String = "A1:AZ200"

' A scope value outside the four the class routes on. It used to be the
' tables-only member of the scope enum, which was a build stage wearing the
' type of a scope and has moved to AnalysisOutput.
Private Const UNKNOWN_SCOPE As Byte = 9

Private Assert As CustomTest

'@section Fixture helpers
'===============================================================================

'@sub-title Free a ListObject name wherever it is taken in the workbook.
'@details
'A ListObject name is unique across the workbook, so the registry name has to
'be free before a fixture claims it.
'@param tableName String. The ListObject name to free.
Private Sub ReleaseTableName(ByVal tableName As String)
    Dim sh As Worksheet
    Dim idx As Long

    For Each sh In ThisWorkbook.Worksheets
        For idx = sh.ListObjects.Count To 1 Step -1
            If StrComp(sh.ListObjects(idx).Name, tableName, vbTextCompare) = 0 Then
                sh.ListObjects(idx).Unlist
            End If
        Next idx
    Next sh
End Sub

'@sub-title Delete every name that points at one worksheet.
'@details
'A name created through Range.Name is workbook-scoped and outlives a clear, and
'the shared helper sweeps the quoted spelling alone. A sheet name with no space
'refers as "=AnaIdsInput!$B$2", so both spellings are matched here.
'@param sh Worksheet. The worksheet whose names are dropped.
Private Sub ReleaseSheetNames(ByVal sh As Worksheet)
    Dim idx As Long
    Dim definitionText As String

    For idx = sh.Names.Count To 1 Step -1
        sh.Names(idx).Delete
    Next idx

    For idx = ThisWorkbook.Names.Count To 1 Step -1
        definitionText = vbNullString
        'A broken name raises on RefersTo. The trap covers this one read.
        On Error Resume Next
        definitionText = ThisWorkbook.Names(idx).RefersTo
        On Error GoTo 0

        If InStr(1, definitionText, "'" & sh.Name & "'!", vbTextCompare) > 0 Or _
           InStr(1, definitionText, "=" & sh.Name & "!", vbTextCompare) > 0 Then
            ThisWorkbook.Names(idx).Delete
        End If
    Next idx
End Sub

'@sub-title Reset one fixture sheet inside a bounded block.
'@param sh Worksheet. The sheet to reset.
Private Sub ResetSheet(ByVal sh As Worksheet)
    Dim idx As Long

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Delete
    Next idx

    For idx = sh.Shapes.Count To 1 Step -1
        sh.Shapes(idx).Delete
    Next idx

    ReleaseSheetNames sh
    sh.Range(FIXTURE_BLOCK).Clear
End Sub

'@sub-title Reset the three fixture sheets and rebuild the registry table.
'@return Worksheet. The registry sheet, carrying an empty registry table.
Private Function ResetFixtures() As Worksheet
    Dim sh As Worksheet

    ResetSheet RegistrySheet()
    ResetSheet InputSheet()
    ResetSheet OutputSheet()

    ReleaseTableName REGISTRY_TABLE
    Set sh = RegistrySheet()
    AnaTabIds.PrepareSheet sh

    Set ResetFixtures = sh
End Function

'@sub-title The sheet hosting the registry table.
Private Function RegistrySheet() As Worksheet
    Set RegistrySheet = ThisWorkbook.Worksheets(REGISTRY_SHEET)
End Function

'@sub-title The sheet standing in for an analysis sheet.
Private Function InputSheet() As Worksheet
    Set InputSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
End Function

'@sub-title The sheet standing in for the sheet the export writes.
Private Function OutputSheet() As Worksheet
    Set OutputSheet = ThisWorkbook.Worksheets(OUTPUT_SHEET)
End Function

'@sub-title Put values and worksheet-scoped names on the input sheet.
'@details
'Four values and their labels, named the way a built analysis table names them.
'The output sheet gets the same values at the same addresses, which is what the
'values-and-formats copy of a real export puts there.
Private Sub BuildInputNames()
    Dim inpsh As Worksheet
    Dim outsh As Worksheet
    Dim counter As Long

    Set inpsh = InputSheet()
    Set outsh = OutputSheet()

    For counter = 1 To 4
        inpsh.Cells(counter + 1, 1).Value = "cat" & counter
        inpsh.Cells(counter + 1, 2).Value = counter * 10
        outsh.Cells(counter + 1, 1).Value = "cat" & counter
        outsh.Cells(counter + 1, 2).Value = counter * 10
    Next

    inpsh.Names.Add Name:="SER_1", _
                    RefersTo:="='" & inpsh.Name & "'!$B$2:$B$5"
    inpsh.Names.Add Name:="LBL_1", _
                    RefersTo:="='" & inpsh.Name & "'!$A$2:$A$5"
End Sub

'@sub-title Register one series row on the fixture registry.
'@param ids AnaTabIds. The instance under test.
'@param scope Byte. The scope to register under.
'@param graphId String. The chart the series belongs to.
'@param address String. Where the chart goes on the output sheet.
Private Sub RegisterSeries(ByVal ids As AnaTabIds, ByVal scope As Byte, _
                           ByVal graphId As String, ByVal address As String)
    ids.AddGraphInfo scope:=scope, graphId:=graphId, _
                     seriesName:="SER_1", seriesType:="bar", _
                     seriesPos:=vbNullString, seriesLabel:="LBL_1", _
                     seriesColumnLabel:=vbNullString, hardCodeLabels:=True, _
                     outRangeAddress:=address
End Sub

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the output sheet, the three fixture sheets and the harness
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    EnsureWorksheet REGISTRY_SHEET, clearSheet:=True, visibility:=xlSheetHidden
    EnsureWorksheet INPUT_SHEET, clearSheet:=True, visibility:=xlSheetHidden
    'WriteGraphs calls Application.GoTo, which needs a selectable sheet.
    EnsureWorksheet OUTPUT_SHEET, clearSheet:=True, visibility:=xlSheetVisible
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnaTabIds"
End Sub

'@sub-title Print results and tear down the fixture sheets
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    ReleaseTableName REGISTRY_TABLE
    ReleaseSheetNames InputSheet()
    ReleaseSheetNames OutputSheet()
    DeleteWorksheets REGISTRY_SHEET, INPUT_SHEET, OUTPUT_SHEET
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush pending assertions after each test
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory and registry building
'===============================================================================

'@sub-title Verify Create rejects a Nothing worksheet
'@TestMethod("AnaTabIds")
Public Sub TestCreateRefusesANothingWorksheet()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateRefusesANothingWorksheet"
    On Error GoTo TestFail

    Dim ids As AnaTabIds

    On Error Resume Next
    Set ids = AnaTabIds.Create(Nothing)
    On Error GoTo 0

    Assert.IsTrue (ids Is Nothing), _
                  "Create with Nothing worksheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRefusesANothingWorksheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create with validation passes on a prepared sheet
'@TestMethod("AnaTabIds")
Public Sub TestCreateWithCheckPasses()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateWithCheckPasses"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)

    Assert.IsTrue (Not ids Is Nothing), _
                  "Create on a prepared sheet should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWithCheckPasses", Err.Number, Err.Description
End Sub

'@sub-title Verify Create with validation refuses a sheet with no registry
'@details
'The description of a raise does not survive the class boundary, so the number
'is the assertion and the text is carried into the message.
'@TestMethod("AnaTabIds")
Public Sub TestCreateWithCheckRefusesASheetWithNoRegistry()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateWithCheckRefusesASheetWithNoRegistry"
    On Error GoTo TestFail

    Dim ids As AnaTabIds
    Dim errNumber As Long
    Dim errText As String

    ResetFixtures
    ReleaseTableName REGISTRY_TABLE

    On Error Resume Next
    Set ids = AnaTabIds.Create(RegistrySheet(), check:=True)
    errNumber = Err.Number
    errText = Err.Description
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A sheet with no registry table should raise InvalidArgument" & _
                    " - description was [" & errText & "]"
    Assert.IsTrue (ids Is Nothing), _
                  "A refused Create should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWithCheckRefusesASheetWithNoRegistry", Err.Number, Err.Description
End Sub

'@sub-title Verify Create without validation accepts any worksheet
'@TestMethod("AnaTabIds")
Public Sub TestCreateWithoutCheck()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateWithoutCheck"
    On Error GoTo TestFail

    Dim ids As AnaTabIds

    ResetFixtures
    Set ids = AnaTabIds.Create(InputSheet(), check:=False)

    Assert.IsTrue (Not ids Is Nothing), _
                  "Create without check should succeed on any worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWithoutCheck", Err.Number, Err.Description
End Sub

'@sub-title Verify PrepareSheet builds one registry table with its headers
'@TestMethod("AnaTabIds")
Public Sub TestPrepareSheetBuildsOneRegistryTable()
    CustomTestSetTitles Assert, "AnaTabIds", "TestPrepareSheetBuildsOneRegistryTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject

    Set sh = ResetFixtures()
    Set lo = sh.ListObjects(REGISTRY_TABLE)

    Assert.AreEqual 1, sh.ListObjects.Count, _
                    "The tracking sheet should carry one table"
    Assert.AreEqual 16, lo.ListColumns.Count, _
                    "The registry should carry every column the class reads"
    Assert.AreEqual "scope", CStr(lo.HeaderRowRange.Cells(1, 1).Value), _
                    "The first column should be the scope"
    Assert.AreEqual "graphId", CStr(lo.HeaderRowRange.Cells(1, 3).Value), _
                    "The third column should be the graph identifier"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPrepareSheetBuildsOneRegistryTable", Err.Number, Err.Description
End Sub

'@sub-title Verify PrepareSheet leaves a table it already built alone
'@TestMethod("AnaTabIds")
Public Sub TestPrepareSheetLeavesAnExistingTableAlone()
    CustomTestSetTitles Assert, "AnaTabIds", "TestPrepareSheetLeavesAnExistingTableAlone"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)
    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"

    AnaTabIds.PrepareSheet sh

    Assert.AreEqual 1, sh.ListObjects.Count, _
                    "A second prepare should leave one table"
    Assert.AreEqual 1, sh.ListObjects(REGISTRY_TABLE).ListRows.Count, _
                    "A second prepare should leave the rows alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPrepareSheetLeavesAnExistingTableAlone", Err.Number, Err.Description
End Sub

'@section Registering rows
'===============================================================================

'@sub-title Verify the first series fills the blank row and the next appends
'@TestMethod("AnaTabIds")
Public Sub TestAddGraphInfoFillsTheFirstRowThenAppends()
    CustomTestSetTitles Assert, "AnaTabIds", "TestAddGraphInfoFillsTheFirstRowThenAppends"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)
    Set lo = sh.ListObjects(REGISTRY_TABLE)

    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"
    Assert.AreEqual 1, lo.ListRows.Count, _
                    "The first series should fill the blank row of a new table"

    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"
    Assert.AreEqual 2, lo.ListRows.Count, _
                    "The second series should be appended"
    Assert.AreEqual "series", CStr(lo.DataBodyRange.Cells(2, 2).Value), _
                    "An appended row should be a series row"
    Assert.AreEqual "G1", CStr(lo.DataBodyRange.Cells(2, 3).Value), _
                    "An appended row should name its chart"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddGraphInfoFillsTheFirstRowThenAppends", Err.Number, Err.Description
End Sub

'@sub-title Verify a series row stores every field in its own column
'@TestMethod("AnaTabIds")
Public Sub TestAddGraphInfoStoresTheSeriesFields()
    CustomTestSetTitles Assert, "AnaTabIds", "TestAddGraphInfoStoresTheSeriesFields"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)
    Set lo = sh.ListObjects(REGISTRY_TABLE)

    ids.AddGraphInfo scope:=AnalysisScopeSpatial, graphId:="G7", _
                     seriesName:="SER_1", seriesType:="line", _
                     seriesPos:="right", seriesLabel:="LBL_1", _
                     seriesColumnLabel:="COL_1", hardCodeLabels:=True, _
                     outRangeAddress:="$H$4", prefix:="adm1_", prefixOnly:=True

    Assert.AreEqual CLng(AnalysisScopeSpatial), CLng(lo.DataBodyRange.Cells(1, 1).Value), _
                    "The row should carry the scope it was registered under"
    Assert.AreEqual "line", CStr(lo.DataBodyRange.Cells(1, 5).Value), _
                    "The series type should be stored"
    Assert.AreEqual "$H$4", CStr(lo.DataBodyRange.Cells(1, 10).Value), _
                    "The chart position should be stored"
    Assert.IsTrue (lo.DataBodyRange.Cells(1, 12).Value <> 0), _
                  "A True prefixOnly should be stored as a number Excel keeps"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddGraphInfoStoresTheSeriesFields", Err.Number, Err.Description
End Sub

'@sub-title Verify a format row carries the look of one chart
'@TestMethod("AnaTabIds")
Public Sub TestAddGraphFormatStoresTheFormatRow()
    CustomTestSetTitles Assert, "AnaTabIds", "TestAddGraphFormatStoresTheFormatRow"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)
    Set lo = sh.ListObjects(REGISTRY_TABLE)

    ids.AddGraphFormat scope:=AnalysisScopeNormal, graphId:="G1", _
                       catTitle:="LBL_1", valuesTitle:="Cases", _
                       hardCodeLabels:=False, heightFactor:=3, _
                       plotTitle:="A plot"

    Assert.AreEqual "format", CStr(lo.DataBodyRange.Cells(1, 2).Value), _
                    "The row should be a format row"
    Assert.AreEqual "A plot", CStr(lo.DataBodyRange.Cells(1, 16).Value), _
                    "The plot title should be stored"
    Assert.AreEqual 3, CLng(lo.DataBodyRange.Cells(1, 15).Value), _
                    "The height factor should be stored"
    Assert.AreEqual 0, CLng(lo.DataBodyRange.Cells(1, 9).Value), _
                    "A False flag should be stored as zero"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddGraphFormatStoresTheFormatRow", Err.Number, Err.Description
End Sub

'@sub-title Verify a scope this class cannot route is refused by name
'@TestMethod("AnaTabIds")
Public Sub TestAddGraphInfoRefusesAnUnknownScope()
    CustomTestSetTitles Assert, "AnaTabIds", "TestAddGraphInfoRefusesAnUnknownScope"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim ids As AnaTabIds
    Dim errNumber As Long
    Dim errText As String

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)
    Set lo = sh.ListObjects(REGISTRY_TABLE)

    On Error Resume Next
    RegisterSeries ids, UNKNOWN_SCOPE, "G1", "$D$2"
    errNumber = Err.Number
    errText = Err.Description
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A scope outside the four should raise InvalidArgument" & _
                    " - description was [" & errText & "]"
    Assert.IsTrue IsEmpty(lo.DataBodyRange.Cells(1, 1)), _
                  "A refused scope should leave the registry empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddGraphInfoRefusesAnUnknownScope", Err.Number, Err.Description
End Sub

'@sub-title Verify WriteGraphs refuses a scope it cannot route
'@TestMethod("AnaTabIds")
Public Sub TestWriteGraphsRefusesAnUnknownScope()
    CustomTestSetTitles Assert, "AnaTabIds", "TestWriteGraphsRefusesAnUnknownScope"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim ids As AnaTabIds
    Dim errNumber As Long
    Dim errText As String

    Set sh = ResetFixtures()
    Set ids = AnaTabIds.Create(sh, check:=True)

    On Error Resume Next
    ids.WriteGraphs OutputSheet(), UNKNOWN_SCOPE, InputSheet()
    errNumber = Err.Number
    errText = Err.Description
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "WriteGraphs should refuse a scope outside the four" & _
                    " - description was [" & errText & "]"
    Assert.AreEqual 0, OutputSheet().ChartObjects.Count, _
                    "A refused scope should draw nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWriteGraphsRefusesAnUnknownScope", Err.Number, Err.Description
End Sub

'@section Name transfer
'===============================================================================

'@sub-title Verify every name of the analysis sheet lands on the output sheet
'@details
'This is the regression guard for the storage change: nothing records these
'names any more, so the export finds them by reading the workbook.
'@TestMethod("AnaTabIds")
Public Sub TestTransferNamesRecreatesTheNamesOfTheSheet()
    CustomTestSetTitles Assert, "AnaTabIds", "TestTransferNamesRecreatesTheNamesOfTheSheet"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    BuildInputNames
    Set outsh = OutputSheet()
    'A name of another sheet, to prove the filter
    RegistrySheet().Range("AC1").Name = "ANAIDS_OTHER"

    Set ids = AnaTabIds.Create(sh, check:=True)
    ids.TransferNames InputSheet(), outsh

    Assert.IsTrue NameOnSheet(outsh, "SER_1"), _
                  "The series name should be recreated on the output sheet"
    Assert.AreEqual "$B$2:$B$5", outsh.Range("SER_1").Address, _
                    "The recreated name should hold the same address"
    Assert.IsTrue NameOnSheet(outsh, "LBL_1"), _
                  "The label name should be recreated on the output sheet"
    Assert.IsTrue Not NameOnSheet(outsh, "ANAIDS_OTHER"), _
                  "A name of another worksheet should stay where it is"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferNamesRecreatesTheNamesOfTheSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify a name holding a value does not stop the transfer
'@details
'A hidden name stores its value inside its definition, so it has no range
'behind it. The walk has to step over one.
'@TestMethod("AnaTabIds")
Public Sub TestTransferNamesStepsOverANameHoldingAValue()
    CustomTestSetTitles Assert, "AnaTabIds", "TestTransferNamesStepsOverANameHoldingAValue"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim ids As AnaTabIds
    Dim errNumber As Long

    Set sh = ResetFixtures()
    BuildInputNames
    InputSheet().Names.Add Name:="ANAIDS_CONST", RefersTo:="=""a value"""

    Set ids = AnaTabIds.Create(sh, check:=True)

    On Error Resume Next
    ids.TransferNames InputSheet(), OutputSheet()
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual 0, errNumber, _
                    "A name holding a value should not stop the transfer"
    Assert.IsTrue NameOnSheet(OutputSheet(), "SER_1"), _
                  "The names with a range behind them should still be carried"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferNamesStepsOverANameHoldingAValue", Err.Number, Err.Description
End Sub

'@section Writing the charts
'===============================================================================

'@sub-title Verify one chart is drawn per graph identifier
'@TestMethod("AnaTabIds")
Public Sub TestWriteGraphsCreatesOneChartPerGraphId()
    CustomTestSetTitles Assert, "AnaTabIds", "TestWriteGraphsCreatesOneChartPerGraphId"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    BuildInputNames
    Set outsh = OutputSheet()

    Set ids = AnaTabIds.Create(sh, check:=True)
    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"
    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"
    RegisterSeries ids, AnalysisScopeNormal, "G2", "$D$30"
    RegisterSeries ids, AnalysisScopeNormal, "G2", "$D$30"
    ids.AddGraphFormat scope:=AnalysisScopeNormal, graphId:="G1", _
                       catTitle:="LBL_1", valuesTitle:="Cases", _
                       hardCodeLabels:=True, plotTitle:="First"
    ids.AddGraphFormat scope:=AnalysisScopeNormal, graphId:="G2", _
                       catTitle:="LBL_1", valuesTitle:="Cases", _
                       hardCodeLabels:=True, plotTitle:="Second"

    ids.WriteGraphs outsh, AnalysisScopeNormal, InputSheet()

    Assert.AreEqual 2, outsh.ChartObjects.Count, _
                    "Two graph identifiers over four series should draw two charts"
    Assert.IsTrue NameOnSheet(outsh, "SER_1"), _
                  "WriteGraphs should carry the names before it draws"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWriteGraphsCreatesOneChartPerGraphId", Err.Number, Err.Description
End Sub

'@sub-title Verify the rows of another scope are left alone
'@TestMethod("AnaTabIds")
Public Sub TestWriteGraphsReadsOneScopeOnly()
    CustomTestSetTitles Assert, "AnaTabIds", "TestWriteGraphsReadsOneScopeOnly"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    BuildInputNames
    Set outsh = OutputSheet()

    Set ids = AnaTabIds.Create(sh, check:=True)
    RegisterSeries ids, AnalysisScopeNormal, "G1", "$D$2"
    RegisterSeries ids, AnalysisScopeSpatial, "G2", "$D$30"
    RegisterSeries ids, AnalysisScopeSpatioTemporal, "G3", "$D$60"

    ids.WriteGraphs outsh, AnalysisScopeNormal, InputSheet()

    Assert.AreEqual 1, outsh.ChartObjects.Count, _
                    "Only the rows of the scope asked for should be drawn"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWriteGraphsReadsOneScopeOnly", Err.Number, Err.Description
End Sub

'@sub-title Verify an empty registry still carries the names
'@TestMethod("AnaTabIds")
Public Sub TestWriteGraphsCarriesTheNamesWithNoChartRow()
    CustomTestSetTitles Assert, "AnaTabIds", "TestWriteGraphsCarriesTheNamesWithNoChartRow"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim ids As AnaTabIds

    Set sh = ResetFixtures()
    BuildInputNames
    Set outsh = OutputSheet()

    Set ids = AnaTabIds.Create(sh, check:=True)
    ids.WriteGraphs outsh, AnalysisScopeSpatial, InputSheet()

    Assert.IsTrue NameOnSheet(outsh, "SER_1"), _
                  "A scope with no chart row should still carry the names"
    Assert.AreEqual 0, outsh.ChartObjects.Count, _
                    "A scope with no chart row should draw nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWriteGraphsCarriesTheNamesWithNoChartRow", Err.Number, Err.Description
End Sub

'@section Assertion helpers
'===============================================================================

'@sub-title Test whether a name resolves to a range of one worksheet.
'@param sh Worksheet. The worksheet the name should point at.
'@param rngName String. The name to look for.
'@return Boolean. True when the name resolves to a range of that sheet.
Private Function NameOnSheet(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range

    'A missing name raises. The trap covers this one read.
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rngName).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then Exit Function
    NameOnSheet = (rng.Worksheet.Name = sh.Name)
End Function
