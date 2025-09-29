Attribute VB_Name = "TestGraphSpecsCache"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the GraphSpecsCache helper class")

Private Const CACHE_SHEET As String = "GraphSpecsCache"

Private Assert As Object
Private Cache As IGraphSpecsCache
Private GraphList As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet CACHE_SHEET
    Set Cache = Nothing
    Set GraphList = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    SeedGraphWorksheet
    Set Cache = GraphSpecsCache.Create(GraphList)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Cache = Nothing
    Set GraphList = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("GraphSpecsCache")
Private Sub TestGraphIdsReturnsUniqueIdentifiers()
    Dim ids As BetterArray
    Dim expected() As String
    Dim idx As Long

    On Error GoTo Fail

    Set ids = Cache.GraphIds
    expected = Array("GraphA", "GraphB")

    Assert.AreEqual 2&, ids.Length, "GraphIds should expose unique identifiers"

    For idx = LBound(expected) To UBound(expected)
        Assert.IsTrue ids.Includes(expected(idx)), "Missing graph id: " & expected(idx)
    Next idx
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestGraphIdsReturnsUniqueIdentifiers"
End Sub

'@TestMethod("GraphSpecsCache")
Private Sub TestColumnValuesReturnsCachedData()
    Dim values As BetterArray

    On Error GoTo Fail

    Set values = Cache.ColumnValues("GraphA", "series id")
    Assert.AreEqual 2&, values.Length
    Assert.AreEqual "Series1", CStr(values.Item(values.LowerBound))
    Assert.AreEqual "Series2", CStr(values.Item(values.LowerBound + 1))

    'ensure repeated calls reuse cache without error
    Set values = Cache.ColumnValues("GraphB", "axis")
    Assert.AreEqual "secondary", NormalizeValue(CStr(values.Item(values.LowerBound)))
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestColumnValuesReturnsCachedData"
End Sub

'@TestMethod("GraphSpecsCache")
Private Sub TestRefreshReflectsListChanges()
    Dim values As BetterArray

    On Error GoTo Fail

    GraphList.ListColumns("axis").DataBodyRange.Cells(3, 1).Value = "secondary"
    Cache.Refresh

    Set values = Cache.ColumnValues("GraphA", "axis")
    Assert.AreEqual 2&, values.Length
    Assert.AreEqual "secondary", NormalizeValue(CStr(values.Item(values.UpperBound)))
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRefreshReflectsListChanges"
End Sub

'@TestMethod("GraphSpecsCache")
Private Sub TestColumnValuesMissingGraphReturnsEmpty()
    Dim values As BetterArray

    On Error GoTo Fail

    Set values = Cache.ColumnValues("Unknown", "series id")
    Assert.AreEqual 0&, values.Length
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestColumnValuesMissingGraphReturnsEmpty"
End Sub

'@section Helpers
'===============================================================================

Private Sub SeedGraphWorksheet()
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(CACHE_SHEET)
    sh.Cells.Clear

    sh.Range("A1").Value = "graph id"
    sh.Range("B1").Value = "series id"
    sh.Range("C1").Value = "axis"
    sh.Range("D1").Value = "type"
    sh.Range("E1").Value = "label"

    sh.Range("A2:E4").Value = Array( _
        Array("GraphA", "Series1", "primary", "bar", "Cases"), _
        Array("GraphA", "Series2", "primary", "line", "Deaths"), _
        Array("GraphB", "Series3", "secondary", "line", "Admissions"))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete

    Set dataRange = sh.Range("A1").CurrentRegion
    Set GraphList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Sub

Private Function NormalizeValue(ByVal valueText As String) As String
    NormalizeValue = LCase$(Trim$(valueText))
End Function

