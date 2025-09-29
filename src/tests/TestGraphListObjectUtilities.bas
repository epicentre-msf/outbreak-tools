Attribute VB_Name = "TestGraphListObjectUtilities"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests validating GraphListObjectUtilities helper functions")

Private Const UTIL_SHEET As String = "GraphUtilSheet"

Private Assert As Object
Private HostList As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet UTIL_SHEET
    Set HostList = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    SeedHostList
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set HostList = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("GraphListObjectUtilities")
Private Sub TestRangeContainsValue()
    Dim headerRange As Range

    On Error GoTo Fail

    Set headerRange = HostList.HeaderRowRange
    Assert.IsTrue RangeContainsValue(headerRange, "Graph ID"), _
        "Should find header regardless of casing"
    Assert.IsFalse RangeContainsValue(headerRange, "Graph ID", True), _
        "Strict search fails due to case mismatch"
    Assert.IsTrue RangeContainsValue(headerRange, "series id", True), _
        "Strict search succeeds when exact case provided"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRangeContainsValue"
End Sub

'@TestMethod("GraphListObjectUtilities")
Private Sub TestListObjectColumnIndex()
    Dim relativeIndex As Long
    Dim absoluteIndex As Long

    On Error GoTo Fail

    relativeIndex = ListObjectColumnIndex(HostList, "axis")
    absoluteIndex = ListObjectColumnIndex(HostList, "axis", False)

    Assert.AreEqual 3&, relativeIndex
    Assert.AreEqual HostList.HeaderRowRange.Cells(1, 3).Column, absoluteIndex
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestListObjectColumnIndex"
End Sub

'@TestMethod("GraphListObjectUtilities")
Private Sub TestUniqueColumnValues()
    Dim values As BetterArray

    On Error GoTo Fail

    Set values = ListObjectUniqueColumnValues(HostList, "graph id")

    Assert.AreEqual 2&, values.Length
    Assert.IsTrue values.Includes("GraphA")
    Assert.IsTrue values.Includes("GraphB")
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestUniqueColumnValues"
End Sub

'@section Helpers
'===============================================================================

Private Sub SeedHostList()
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(UTIL_SHEET)
    sh.Cells.Clear

    sh.Range("A1").Value = "graph id"
    sh.Range("B1").Value = "series id"
    sh.Range("C1").Value = "axis"

    sh.Range("A2:C4").Value = Array( _
        Array("GraphA", "Series1", "primary"), _
        Array("GraphA", "Series2", "primary"), _
        Array("GraphB", "Series3", "secondary"))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete

    Set dataRange = sh.Range("A1").CurrentRegion
    Set HostList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Sub

