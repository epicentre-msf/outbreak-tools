Attribute VB_Name = "TestSetupTranslationsTable"
Attribute VB_Description = "Unit tests for the improved translations table manager"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the SetupTranslationsTable class covering caching, registry updates and language management")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private TranslationsSheet As Worksheet
Private RegistrySheet As Worksheet
Private SourceSheet As Worksheet
Private TranslationsTable As ListObject
Private RegistryTable As ListObject
Private Subject As ISetupTranslationsTable

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const REGISTRY_SHEET_NAME As String = "Registry"
Private Const SOURCE_SHEET_NAME As String = "SourceData"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const REGISTRY_TABLE_NAME As String = "Tab_Registry"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"

'@ModuleInitialize
Private Sub ModuleInitialize()
    TestHelpers.BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupTranslationsTable"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set TranslationsSheet = TestHelpers.EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)
    Set RegistrySheet = TestHelpers.EnsureWorksheet(REGISTRY_SHEET_NAME, FixtureWorkbook)
    Set SourceSheet = TestHelpers.EnsureWorksheet(SOURCE_SHEET_NAME, FixtureWorkbook)

    Set TranslationsTable = BuildTranslationsTable(TranslationsSheet)
    Set RegistryTable = BuildRegistryTable(RegistrySheet)
    RegisterSourceRanges SourceSheet, FixtureWorkbook

    Set Subject = SetupTranslationsTable.Create(TranslationsTable)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
        FixtureWorkbook = Nothing
    On Error GoTo 0

    Set Subject = Nothing
    Set RegistryTable = Nothing
    Set TranslationsTable = Nothing
    Set SourceSheet = Nothing
    Set RegistrySheet = Nothing
    Set TranslationsSheet = Nothing
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestCreateRejectsMissingTable()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestCreateRejectsMissingTable"

    On Error GoTo ExpectError
        Dim invalid As ISetupTranslationsTable
        Set invalid = SetupTranslationsTable.Create(Nothing)
        Assert.LogFailure "Create should reject a missing listobject"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Create must raise InvalidArgument when the listobject is missing"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEnsureLanguagesAddsUniqueColumns()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEnsureLanguagesAddsUniqueColumns"

    Subject.EnsureLanguages "French;French;German;"

    Assert.AreEqual CLng(3), TranslationsTable.ListColumns.Count, "Should add two extra language columns without duplicates"
    Assert.IsTrue HasColumn("English"), "Existing base column should remain"
    Assert.IsTrue HasColumn("French"), "French column should be created"
    Assert.IsTrue HasColumn("German"), "German column should be created"
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryAddsLabelsAndTags()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryAddsLabelsAndTags"

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "Six unique labels expected after processing text and formula ranges"
    Assert.AreEqual "RNG_Greetings--1", TagForLabel("Hello"), "Existing labels should reuse the helper column tag"
    Assert.AreEqual "RNG_Greetings--1", TagForLabel("Good bye"), "Second entry from greetings range should be tagged accordingly"
    Assert.AreEqual "RNG_Farewell--1", TagForLabel("Farewell"), "Farewell range should be imported on first execution even with status no"
    Assert.AreEqual "RNG_Formula--1", TagForLabel("Morning"), "Formula text Morning should be extracted and tagged"
    Assert.IsTrue WorkbookHasName(COUNTER_NAME), "Update sequence counter should be stored as a workbook name"
    Assert.AreEqual CLng(1), FixtureWorkbook.Names(COUNTER_NAME).RefersToRange.Value, "Counter should be incremented to one after first update"
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistrySkipsWhenStatusNo()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistrySkipsWhenStatusNo"

    Subject.UpdateFromRegistry RegistrySheet
    SetRegistryStatus "yes", "no", "no"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "No additional rows should be created when statuses are no"
    Assert.AreEqual "RNG_Greetings--2", TagForLabel("Hello"), "Existing label should update tag with the new sequence number"
    Assert.AreEqual CLng(2), FixtureWorkbook.Names(COUNTER_NAME).RefersToRange.Value, "Counter must be incremented to two after the second update"
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryRejectsUnknownMode()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryRejectsUnknownMode"

    RegistryTable.ListRows(1).Range.Cells(1, 3).Value = "unsupported"

    On Error GoTo ExpectError
        Subject.UpdateFromRegistry RegistrySheet
        Assert.LogFailure "Unsupported translation mode should raise an error"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Invalid translation mode must raise InvalidArgument"
    Err.Clear
End Sub

'@section Helpers
'===============================================================================
Private Sub AssertSheetSetup()
    TestHelpers.EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
End Sub

Private Function BuildTranslationsTable(ByVal targetSheet As Worksheet) As ListObject
    targetSheet.Cells.Clear
    targetSheet.Cells(1, 1).Value = "TranslationTag"
    targetSheet.Cells(1, 2).Value = "English"

    Dim tableRange As Range
    Set tableRange = targetSheet.Range("B1:B2")

    Dim table As ListObject
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = TRANSLATIONS_TABLE_NAME

    Set BuildTranslationsTable = table
End Function

Private Function BuildRegistryTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    matrix = TestHelpers.RowsToMatrix(Array( _
        Array("rngname", "status", "mode"), _
        Array("RNG_Greetings", "yes", "translate as text"), _
        Array("RNG_Farewell", "no", "translate as text"), _
        Array("RNG_Formula", "yes", "translate as formula")))

    targetSheet.Cells.Clear
    TestHelpers.WriteMatrix targetSheet.Cells(1, 1), matrix

    Dim registryRange As Range
    Set registryRange = targetSheet.Range("A1:C4")

    Dim table As ListObject
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=registryRange, XlListObjectHasHeaders:=xlYes)
    table.Name = REGISTRY_TABLE_NAME

    Set BuildRegistryTable = table
End Function

Private Sub RegisterSourceRanges(ByVal targetSheet As Worksheet, ByVal hostWorkbook As Workbook)
    targetSheet.Cells.Clear

    targetSheet.Range("A1").Value = "Hello"
    targetSheet.Range("A2").Value = "Good bye"
    targetSheet.Range("B1").Value = "Farewell"
    targetSheet.Range("B2").Value = "See you"
    targetSheet.Range("C1").Formula = "=IF(A1="""", ""Morning"", ""Evening"")"

    hostWorkbook.Names.Add Name:="RNG_Greetings", RefersTo:=targetSheet.Range("A1:A2")
    hostWorkbook.Names.Add Name:="RNG_Farewell", RefersTo:=targetSheet.Range("B1:B2")
    hostWorkbook.Names.Add Name:="RNG_Formula", RefersTo:=targetSheet.Range("C1")
End Sub

Private Sub SetRegistryStatus(ByVal firstStatus As String, ByVal secondStatus As String, ByVal thirdStatus As String)
    RegistryTable.ListRows(1).Range.Cells(1, 2).Value = firstStatus
    RegistryTable.ListRows(2).Range.Cells(1, 2).Value = secondStatus
    RegistryTable.ListRows(3).Range.Cells(1, 2).Value = thirdStatus
End Sub

Private Function TagForLabel(ByVal label As String) As String
    Dim row As ListRow

    For Each row In TranslationsTable.ListRows
        If StrComp(CStr(row.Range.Cells(1, 1).Value), label, vbBinaryCompare) = 0 Then
            TagForLabel = CStr(row.Range.Cells(1, 1).Offset(0, -1).Value)
            Exit Function
        End If
    Next row
End Function

Private Function HasColumn(ByVal columnName As String) As Boolean
    Dim column As ListColumn
    For Each column In TranslationsTable.ListColumns
        If StrComp(column.Name, columnName, vbTextCompare) = 0 Then
            HasColumn = True
            Exit Function
        End If
    Next column
End Function

Private Function WorkbookHasName(ByVal nameText As String) As Boolean
    Dim definedName As Name
    On Error Resume Next
        Set definedName = FixtureWorkbook.Names(nameText)
    On Error GoTo 0
    WorkbookHasName = Not (definedName Is Nothing)
End Function
