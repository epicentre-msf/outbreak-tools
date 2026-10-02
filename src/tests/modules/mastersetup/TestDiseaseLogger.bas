Attribute VB_Name = "TestDiseaseLogger"
Attribute VB_Description = "Tests covering DiseaseLogger behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering DiseaseLogger behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest
Private Logger As IDiseaseLogger

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseLogger"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Logger = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Logger = New DiseaseLogger
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Logger = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseLogger")
Public Sub TestRecordAddsEntries()
    CustomTestSetTitles Assert, "DiseaseLogger", "TestRecordAddsEntries"

    Dim details As BetterArray
    Set details = New BetterArray
    details.Push "Variable"
    details.Push "var_age"

    Logger.Record "Import", DiseaseLogInfo, "Appended variable", details

    Assert.IsTrue Logger.HasEntries, "Logger should report entries after Record"
    Dim entries As BetterArray
    Set entries = Logger.Entries
    Assert.AreEqual 1, entries.Length, "Entries should contain one record"

    Dim firstEntry As BetterArray
    Set firstEntry = entries.Item(entries.LowerBound)
    Assert.AreEqual "Import", firstEntry.Item(firstEntry.LowerBound + 1), "Operation should be preserved"
    Assert.AreEqual "Appended variable", firstEntry.Item(firstEntry.LowerBound + 3), "Message should be preserved"
End Sub

'@TestMethod("DiseaseLogger")
Public Sub TestClearRemovesEntries()
    CustomTestSetTitles Assert, "DiseaseLogger", "TestClearRemovesEntries"

    Logger.Record "Export", DiseaseLogWarning, "Skipped disease"

    Logger.Clear
    Assert.IsFalse Logger.HasEntries, "Logger should be empty after Clear"
    Assert.AreEqual 0, Logger.Entries.Length, "Entries collection should be empty after Clear"
End Sub

'@TestMethod("DiseaseLogger")
Public Sub TestRecordRequiresOperation()
    CustomTestSetTitles Assert, "DiseaseLogger", "TestRecordRequiresOperation"

    Dim raisedError As Boolean

    On Error Resume Next
        Logger.Record vbNullString, DiseaseLogInfo, "Missing operation"
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Logger should validate operation"
End Sub
