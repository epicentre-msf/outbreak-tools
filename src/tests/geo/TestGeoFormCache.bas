Attribute VB_Name = "TestGeoFormCache"
Attribute VB_Description = "Tests for GeoFormCache class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for GeoFormCache class")

'@description
'Validates the GeoFormCache class, which holds the four lists the geo form
'searches, read once per load from their workbook names and held in memory.
'
'EVERY FIXTURE IS A WORKBOOK OF ITS OWN
'-------------------------------------------------------------------------------
'BuildGeoWorkbook answers a fresh workbook whose first sheet carries the four
'columns, with the workbook scoped names the production linelist carries:
'adm4_concat, hf_concat, histo_geo and histo_hf. Names added to this workbook
'would collide with the fixtures of the other geo suites, and a workbook is
'deleted whole, names included.
'
'THE TESTS BUILD THEIR OWN INSTANCES
'-------------------------------------------------------------------------------
'Create answers a separate instance per test, so the shared default instance
'the form reads is never touched here and holds nothing when the suite ends.
'
'THE CONCAT COLUMNS ARE WRITTEN OUT OF ORDER ON PURPOSE
'-------------------------------------------------------------------------------
'P2 comes before P1, so a sorted answer proves the load-time sort ran and an
'answer in table order proves the historic lists were left alone.
'@depends GeoFormCache, CustomTest, TestHelpersLite, BetterArray

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Suppresses screen updates via BusyApp, ensures the test output sheet exists,
'creates the CustomTest assertion object targeting that sheet, and sets the
'module name for result grouping.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGeoFormCache"
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, restores the
'application state via RestoreApp, and releases the assertion object. The
'fixtures are workbooks deleted by their own tests, so nothing is owed here.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Reset state before each individual test.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Clean up after each individual test.
'@details
'Flushes any pending assertion results to the output sheet so each test's
'outcome is recorded before the next test begins.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Fixture helpers
'===============================================================================

'@sub-title Build a workbook carrying the four geo form names.
'@details
'The first sheet holds one column per list, and each column is covered by the
'workbook scoped name the production linelist writes. The concat columns are
'out of order on purpose. When withHFNames is False the two facility names
'are left out, which is what the missing-name test reads against.
'@param withHFNames Optional Boolean. True to add hf_concat and histo_hf. Defaults to True.
'@return Workbook. The fixture workbook.
Private Function BuildGeoWorkbook(Optional ByVal withHFNames As Boolean = True) As Workbook
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = NewWorkbook()
    Set sh = wb.Worksheets(1)

    WriteColumn sh.Cells(1, 1), "P2 | D1", "P1 | D1", "P3 | D1"
    wb.Names.Add Name:="adm4_concat", _
                 RefersTo:=sh.Range(sh.Cells(1, 1), sh.Cells(3, 1))

    WriteColumn sh.Cells(1, 3), "P3 | D1", "P1 | D1"
    wb.Names.Add Name:="histo_geo", _
                 RefersTo:=sh.Range(sh.Cells(1, 3), sh.Cells(2, 3))

    If withHFNames Then
        WriteColumn sh.Cells(1, 2), "HF B", "HF A"
        wb.Names.Add Name:="hf_concat", _
                     RefersTo:=sh.Range(sh.Cells(1, 2), sh.Cells(2, 2))

        WriteColumn sh.Cells(1, 4), "HF Z", "HF A"
        wb.Names.Add Name:="histo_hf", _
                     RefersTo:=sh.Range(sh.Cells(1, 4), sh.Cells(2, 4))
    End If

    Set BuildGeoWorkbook = wb
End Function

'@sub-title Read one entry of a list by its position.
'@param valuesTable BetterArray. The list answered by the cache.
'@param position Long. The entry wanted, 1 for the first.
'@return String. The value held there.
Private Function ItemAt(ByVal valuesTable As BetterArray, _
                        ByVal position As Long) As String
    ItemAt = CStr(valuesTable.Item(valuesTable.LowerBound + position - 1))
End Function

'@section Factory tests
'===============================================================================

'@sub-title Verify Create refuses a Nothing workbook with ObjectNotInitialized.
'@details
'Acts by calling GeoFormCache.Create with Nothing under On Error Resume Next
'and captures the error number. Asserts the number is ObjectNotInitialized and
'the result is Nothing.
'@TestMethod("GeoFormCache")
Public Sub TestCreateRejectsNothing()
    CustomTestSetTitles Assert, "GeoFormCache", "TestCreateRejectsNothing"
    On Error GoTo TestFail

    Dim cache As GeoFormCache
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
    Set cache = GeoFormCache.Create(Nothing)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Create with a Nothing workbook should raise ObjectNotInitialized - " & _
                    "description was [" & errDescription & "]"
    Assert.IsTrue (cache Is Nothing), _
                  "Create with a Nothing workbook should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify Create answers a loaded instance.
'@details
'Arranges a fixture workbook. Acts by calling Create and asking for the admin
'concat list. Asserts the instance is there and the list holds the three
'places of the column.
'@TestMethod("GeoFormCache")
Public Sub TestCreateAnswersALoadedInstance()
    CustomTestSetTitles Assert, "GeoFormCache", "TestCreateAnswersALoadedInstance"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)

    Assert.IsNotNothing cache, "Create with a workbook should succeed"

    Set concatTab = cache.ConcatList(GeoScopeAdmin)
    Assert.AreEqual CLng(3), CLng(concatTab.Length), _
                    "The admin concat list should hold the three places of the column"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateAnswersALoadedInstance", Err.Number, Err.Description
End Sub

'@section Reading tests
'===============================================================================

'@sub-title Verify the concat list comes back sorted.
'@details
'Arranges a fixture whose adm4_concat column reads P2, P1, P3. Acts by asking
'for the admin concat list. Asserts the three entries arrive in order, which
'is the load-time sort the search hits rely on.
'@TestMethod("GeoFormCache")
Public Sub TestConcatListComesBackSorted()
    CustomTestSetTitles Assert, "GeoFormCache", "TestConcatListComesBackSorted"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    Set concatTab = cache.ConcatList(GeoScopeAdmin)

    Assert.AreEqual "P1 | D1", ItemAt(concatTab, 1), _
                    "The first entry should be the alphabetically first place"
    Assert.AreEqual "P2 | D1", ItemAt(concatTab, 2), _
                    "The second entry should follow in order"
    Assert.AreEqual "P3 | D1", ItemAt(concatTab, 3), _
                    "The last entry should be the alphabetically last place"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestConcatListComesBackSorted", Err.Number, Err.Description
End Sub

'@sub-title Verify a list is read once and held across asks.
'@details
'Arranges a fixture, asks for the admin concat list, then rewrites the first
'cell of the column on the sheet. Acts by asking again. Asserts the answer
'still holds the values of the first read, which is the cache doing its job:
'the worksheet is crossed once per load.
'@TestMethod("GeoFormCache")
Public Sub TestConcatListIsHeldAcrossAsks()
    CustomTestSetTitles Assert, "GeoFormCache", "TestConcatListIsHeldAcrossAsks"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    cache.ConcatList GeoScopeAdmin

    wb.Worksheets(1).Cells(1, 1).Value = "ZZ | D1"

    Set concatTab = cache.ConcatList(GeoScopeAdmin)

    Assert.AreEqual "P2 | D1", ItemAt(concatTab, 2), _
                    "The held list should keep the value read before the sheet changed"
    Assert.AreEqual CLng(3), CLng(concatTab.Length), _
                    "The held list should keep its length"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestConcatListIsHeldAcrossAsks", Err.Number, Err.Description
End Sub

'@sub-title Verify LoadFrom drops the held lists.
'@details
'Arranges a loaded cache, rewrites a cell of the concat column, and acts by
'calling LoadFrom with the same workbook. Asserts the next ask reads the sheet
'fresh, which is what keeps the form current after a geobase import.
'@TestMethod("GeoFormCache")
Public Sub TestLoadFromDropsTheHeldLists()
    CustomTestSetTitles Assert, "GeoFormCache", "TestLoadFromDropsTheHeldLists"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    cache.ConcatList GeoScopeAdmin

    wb.Worksheets(1).Cells(1, 1).Value = "ZZ | D1"
    cache.LoadFrom wb

    Set concatTab = cache.ConcatList(GeoScopeAdmin)

    Assert.AreEqual "ZZ | D1", ItemAt(concatTab, 3), _
                    "A reload should read the rewritten place off the sheet"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLoadFromDropsTheHeldLists", Err.Number, Err.Description
End Sub

'@sub-title Verify Refresh drops the held lists and keeps the workbook.
'@details
'Arranges a loaded cache, rewrites a cell, and acts by calling Refresh.
'Asserts the next ask reads the sheet fresh on Refresh alone.
'@TestMethod("GeoFormCache")
Public Sub TestRefreshDropsTheHeldLists()
    CustomTestSetTitles Assert, "GeoFormCache", "TestRefreshDropsTheHeldLists"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    cache.ConcatList GeoScopeAdmin

    wb.Worksheets(1).Cells(1, 1).Value = "ZZ | D1"
    cache.Refresh

    Set concatTab = cache.ConcatList(GeoScopeAdmin)

    Assert.AreEqual "ZZ | D1", ItemAt(concatTab, 3), _
                    "A refresh should read the rewritten place off the sheet"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRefreshDropsTheHeldLists", Err.Number, Err.Description
End Sub

'@sub-title Verify the historic list keeps its table order.
'@details
'Arranges a fixture whose histo_geo column reads P3 then P1. Acts by asking
'for the admin historic list. Asserts the entries arrive in that order, which
'is the order the user validated places in.
'@TestMethod("GeoFormCache")
Public Sub TestHistoricListKeepsTableOrder()
    CustomTestSetTitles Assert, "GeoFormCache", "TestHistoricListKeepsTableOrder"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim historicTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    Set historicTab = cache.HistoricList(GeoScopeAdmin)

    Assert.AreEqual "P3 | D1", ItemAt(historicTab, 1), _
                    "The first historic entry should be the first table row"
    Assert.AreEqual "P1 | D1", ItemAt(historicTab, 2), _
                    "The second historic entry should be the second table row"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHistoricListKeepsTableOrder", Err.Number, Err.Description
End Sub

'@sub-title Verify the facility scope reads the facility names.
'@details
'Arranges the full fixture. Acts by asking for both facility lists. Asserts
'the concat list arrives sorted from hf_concat and the historic list keeps the
'table order of histo_hf.
'@TestMethod("GeoFormCache")
Public Sub TestTheFacilityScopeReadsTheFacilityNames()
    CustomTestSetTitles Assert, "GeoFormCache", "TestTheFacilityScopeReadsTheFacilityNames"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray
    Dim historicTab As BetterArray

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    Set concatTab = cache.ConcatList(GeoScopeHF)
    Set historicTab = cache.HistoricList(GeoScopeHF)

    Assert.AreEqual "HF A", ItemAt(concatTab, 1), _
                    "The facility concat list should arrive sorted"
    Assert.AreEqual "HF Z", ItemAt(historicTab, 1), _
                    "The facility historic list should keep its table order"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFacilityScopeReadsTheFacilityNames", Err.Number, Err.Description
End Sub

'@section Refusal tests
'===============================================================================

'@sub-title Verify an unknown scope raises InvalidArgument.
'@details
'Arranges a loaded cache. Acts by asking both readers for GeoScopeBoth, which
'has no list in the form. Asserts each ask raises InvalidArgument. An unknown
'scope used to mean facility in silence elsewhere in the geo code.
'@TestMethod("GeoFormCache")
Public Sub TestAnUnknownScopeRaises()
    CustomTestSetTitles Assert, "GeoFormCache", "TestAnUnknownScopeRaises"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatError As Long
    Dim historicError As Long

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)

    On Error Resume Next
    cache.ConcatList GeoScopeBoth
    concatError = Err.Number
    Err.Clear
    cache.HistoricList GeoScopeBoth
    historicError = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), concatError, _
                    "ConcatList should refuse a scope outside the two the form knows"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), historicError, _
                    "HistoricList should refuse a scope outside the two the form knows"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnknownScopeRaises", Err.Number, Err.Description
End Sub

'@sub-title Verify a cleared cache answers empty lists and raises nothing.
'@details
'Arranges a loaded cache and acts by calling Clear, then asking for both admin
'lists under On Error Resume Next. Asserts nothing raised and both lists are
'empty. This is the state the shared instance is in after a reset, where a
'keystroke in the search box has to show nothing and stay quiet.
'@TestMethod("GeoFormCache")
Public Sub TestAClearedCacheAnswersEmptyLists()
    CustomTestSetTitles Assert, "GeoFormCache", "TestAClearedCacheAnswersEmptyLists"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim concatTab As BetterArray
    Dim historicTab As BetterArray
    Dim errNumber As Long
    Dim errDescription As String

    Set wb = BuildGeoWorkbook()
    Set cache = GeoFormCache.Create(wb)
    cache.ConcatList GeoScopeAdmin
    cache.Clear

    On Error Resume Next
    Set concatTab = cache.ConcatList(GeoScopeAdmin)
    Set historicTab = cache.HistoricList(GeoScopeAdmin)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(0), errNumber, _
                    "A cleared cache should raise nothing - " & _
                    "description was [" & errDescription & "]"
    Assert.AreEqual CLng(0), CLng(concatTab.Length), _
                    "A cleared cache should answer an empty concat list"
    Assert.AreEqual CLng(0), CLng(historicTab.Length), _
                    "A cleared cache should answer an empty historic list"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAClearedCacheAnswersEmptyLists", Err.Number, Err.Description
End Sub

'@sub-title Verify a missing workbook name raises in the read.
'@details
'Arranges a fixture without the two facility names. Acts by asking for the
'facility concat list under On Error Resume Next. Asserts something raised. A
'loaded cache with a broken workbook has to say so, and the callers that load
'the form hold a handler for it.
'@TestMethod("GeoFormCache")
Public Sub TestAMissingNameRaisesInTheRead()
    CustomTestSetTitles Assert, "GeoFormCache", "TestAMissingNameRaisesInTheRead"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim cache As GeoFormCache
    Dim errNumber As Long

    Set wb = BuildGeoWorkbook(withHFNames:=False)
    Set cache = GeoFormCache.Create(wb)

    On Error Resume Next
    cache.ConcatList GeoScopeHF
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (errNumber <> 0), _
                  "Asking for a list whose workbook name is missing should raise"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMissingNameRaisesInTheRead", Err.Number, Err.Description
End Sub
