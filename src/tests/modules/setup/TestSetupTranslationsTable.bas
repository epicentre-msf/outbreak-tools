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
Private Const TAG_SEPARATOR As String = "__"
Private Const LANGUAGES_NAME_ID As String = "__SetupTranslationsLanguages__"

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
    Subject.SetDisplayPrompts False
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
    On Error GoTo Fail

    Subject.EnsureLanguages "French;French;German;"

    Assert.AreEqual CLng(3), TranslationsTable.ListColumns.Count, "Should add two extra language columns without duplicates"
    Assert.IsTrue HasColumn("English"), "Existing base column should remain"
    Assert.IsTrue HasColumn("French"), "French column should be created"
    Assert.IsTrue HasColumn("German"), "German column should be created"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureLanguagesAddsUniqueColumns", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestLanguagesListsNonDefaultHeaders()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLanguagesListsNonDefaultHeaders"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German"

    Dim languages As BetterArray
    Set languages = Subject.Languages

    Assert.AreEqual CLng(2), languages.Length, "Languages should contain each non-default header"
    Assert.AreEqual "French", CStr(languages.Item(languages.LowerBound)), "Languages should follow table column order"
    Assert.AreEqual "German", CStr(languages.Item(languages.LowerBound + 1)), "Languages should include subsequent columns"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLanguagesListsNonDefaultHeaders", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEnsureLanguagesPersistsHiddenName()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEnsureLanguagesPersistsHiddenName"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German"

    Dim store As IHiddenNames
    Dim storedValue As String
    Set store = HiddenNames.Create(TranslationsSheet)
    storedValue = store.ValueAsString(LANGUAGES_NAME_ID)

    Assert.AreEqual "English;French;German", storedValue, "Hidden name should store every language including the default"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureLanguagesPersistsHiddenName", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestLanguagesCanIncludeDefaultHeader()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLanguagesCanIncludeDefaultHeader"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German"

    Dim languages As BetterArray
    Set languages = Subject.Languages(True)

    Assert.AreEqual CLng(3), languages.Length, "Languages should include the default column when requested"
    Assert.AreEqual "English", CStr(languages.Item(languages.LowerBound)), "Default header should be listed first"
    Assert.AreEqual "French", CStr(languages.Item(languages.LowerBound + 1)), "Non-default languages should follow in column order"
    Assert.AreEqual "German", CStr(languages.Item(languages.LowerBound + 2)), "All remaining languages should be included"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLanguagesCanIncludeDefaultHeader", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestExportStartsAtSecondColumnAndCopiesHiddenNames()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestExportStartsAtSecondColumnAndCopiesHiddenNames"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim exportedStore As IHiddenNames
    Dim expectedLanguages As String

    On Error GoTo Fail

    Subject.EnsureLanguages "French"
    TranslationsTable.DataBodyRange.Cells(1, 1).Value = "Hello"

    expectedLanguages = HiddenNames.Create(TranslationsSheet).ValueAsString(LANGUAGES_NAME_ID)

    Set exportBook = TestHelpers.NewWorkbook
    Subject.Export exportBook

    Set exportedSheet = exportBook.Worksheets(TRANSLATIONS_SHEET_NAME)
    Assert.AreEqual "english", LCase$(CStr(exportedSheet.Cells(1, 2).Value)), _
                    "Export should write the first header starting on the second column."

    Set exportedStore = HiddenNames.Create(exportedSheet)
    Assert.AreEqual expectedLanguages, exportedStore.ValueAsString(LANGUAGES_NAME_ID), _
                    "Export should copy translation hidden names into the destination workbook."

    exportBook.Close SaveChanges:=False
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportStartsAtSecondColumnAndCopiesHiddenNames", Err.Number, Err.Description
    On Error Resume Next
        If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryAddsLabelsAndTags()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryAddsLabelsAndTags"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "Six unique labels expected after processing text and formula ranges"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Hello"), "Existing labels should reuse the helper column tag"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Good bye"), "Second entry from greetings range should be tagged accordingly"
    Assert.AreEqual ExpectedTag("RNG_Farewell", 1), TagForLabel("Farewell"), "Farewell range should be imported on first execution even with status no"
    Assert.AreEqual ExpectedTag("RNG_Formula", 1), TagForLabel("Morning"), "Formula text Morning should be extracted and tagged"
    Assert.IsTrue HiddenCounterExists(), "Update sequence counter should be stored using the hidden names manager"
    Assert.AreEqual CLng(1), CounterValue(), "Counter should be incremented to one after first update"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryAddsLabelsAndTags", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistrySkipsWhenStatusNo()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistrySkipsWhenStatusNo"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    SetRegistryStatus "yes", "no", "no"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "No additional rows should be created when statuses are no"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 2), TagForLabel("Hello"), "Existing label should update tag with the new sequence number"
    Assert.AreEqual CLng(2), CounterValue(), "Counter must be incremented to two after the second update"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistrySkipsWhenStatusNo", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryRejectsUnknownMode()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryRejectsUnknownMode"

    RegistryTable.ListRows(1).Range.Cells(1, 4).Value = "unsupported"

    On Error GoTo ExpectError
        Subject.UpdateFromRegistry RegistrySheet
        Assert.LogFailure "Unsupported translation mode should raise an error"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Invalid translation mode must raise InvalidArgument"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryRequiresHelperColumn()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryRequiresHelperColumn"
    On Error GoTo ExpectError

    TranslationsSheet.Columns(1).Delete
    Subject.UpdateFromRegistry RegistrySheet

    Assert.LogFailure "UpdateFromRegistry should raise when the helper column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), Err.Number, "Missing helper column must raise ErrorUnexpectedState"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestResetSequenceSetsCounterToZero()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestResetSequenceSetsCounterToZero"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    Subject.ResetSequence RegistrySheet

    Assert.AreEqual CLng(0), CounterValue(), "ResetSequence should reset the workbook counter to zero"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResetSequenceSetsCounterToZero", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryDeletesMissingLabels()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryDeletesMissingLabels"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    SourceSheet.Range("A2").Value = vbNullString
    SetRegistryStatus "yes", "yes", "yes"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(5), TranslationsTable.ListRows.Count, "Removing a label from a processed range should delete the corresponding translation row"
    Assert.AreEqual vbNullString, TagForLabel("Good bye"), "Deleted labels should no longer be present in the translations table"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 2), TagForLabel("Hello"), "Existing labels must be retagged with the current update sequence"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryDeletesMissingLabels", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryMaintainsSortedOrderAfterCacheRebuild()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryMaintainsSortedOrderAfterCacheRebuild"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    SetRegistryStatus "yes", "yes", "yes"
    SourceSheet.Range("A1").Value = "Apple"
    SourceSheet.Range("B2").Value = "Zulu"

    Subject.UpdateFromRegistry RegistrySheet

    Dim labels As Variant
    labels = TranslationsTable.ListColumns("English").DataBodyRange.Value

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "Cache rebuild should keep six translation rows populated"
    Assert.AreEqual "Apple", CStr(labels(1, 1)), "First label should sort alphabetically after rebuild"
    Assert.AreEqual "Evening", CStr(labels(2, 1)), "Formula tokens should remain in sorted order"
    Assert.AreEqual "Farewell", CStr(labels(3, 1)), "Existing labels should remain sorted post-refresh"
    Assert.AreEqual "Good bye", CStr(labels(4, 1)), "Greetings range should continue contributing labels"
    Assert.AreEqual "Morning", CStr(labels(5, 1)), "Formula chunk order should be stable"
    Assert.AreEqual "Zulu", CStr(labels(6, 1)), "Updated farewell range should sort to the bottom"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 2), TagForLabel("Apple"), "Updated greetings label should receive latest sequence tag"
    Assert.AreEqual ExpectedTag("RNG_Farewell", 2), TagForLabel("Zulu"), "Farewell update should advance to the new sequence"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryMaintainsSortedOrderAfterCacheRebuild", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryProcessesSingleCellRegistryTable()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryProcessesSingleCellRegistryTable"
    On Error GoTo Fail

    SourceSheet.Range("D1").Value = "Solo"
    FixtureWorkbook.Names.Add Name:="RNG_Solo", RefersTo:=SourceSheet.Range("D1")

    Dim singleMatrix As Variant
    singleMatrix = TestHelpers.RowsToMatrix(Array( _ 
                                                Array("tabname", "rngname", "status", "mode"), _ 
                                                Array("table", "RNG_Solo", "yes", "translate as text")))
    TestHelpers.WriteMatrix RegistrySheet.Range("F1"), singleMatrix

    Dim singleTable As ListObject
    Set singleTable = RegistrySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=RegistrySheet.Range("F1:I2"), XlListObjectHasHeaders:=xlYes)
    singleTable.Name = "Tab_RegistrySingle"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual ExpectedTag("RNG_Solo", 1), TagForLabel("Solo"), "Single-cell registry watcher should process its named range"
    Assert.AreEqual CLng(7), TranslationsTable.ListRows.Count, "New single-cell watcher should add an extra translation row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryProcessesSingleCellRegistryTable", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestNumberOfMissingReportsPerLanguage()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestNumberOfMissingReportsPerLanguage"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Dim summary As String
    summary = Subject.NumberOfMissing

    Assert.AreEqual "Translation Updated!" & vbLf & "6 labels are missing for column French.", summary, "NumberOfMissing should report missing counts for each non default language"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestNumberOfMissingReportsPerLanguage", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestMissingLabelsCountsBlankCells()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestMissingLabelsCountsBlankCells"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Dim missing As Long
    missing = Subject.MissingLabels("French")

    Assert.AreEqual CLng(TranslationsTable.ListRows.Count), missing, "MissingLabels should count each blank entry in the target language column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingLabelsCountsBlankCells", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestMissingLabelsReturnsZeroWhenTranslationsPresent()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestMissingLabelsReturnsZeroWhenTranslationsPresent"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"
    TranslationsTable.ListColumns("French").DataBodyRange.Value = "french-text"

    Assert.AreEqual CLng(0), Subject.MissingLabels("French"), "MissingLabels should return zero when the language column has no blanks"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingLabelsReturnsZeroWhenTranslationsPresent", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateLabelsReturnsEmptyWhenAllLabelsUnique()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateLabelsReturnsEmptyWhenAllLabelsUnique"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTranslationLabel "Alpha"
    AppendTranslationLabel "Beta"
    AppendTranslationLabel "Gamma"

    Dim summary As String
    Assert.IsFalse Subject.DuplicateLabels(summary), "DuplicateLabels should return False when no duplicates exist"
    Assert.AreEqual vbNullString, summary, "DuplicateLabels should not populate the message when no duplicates exist"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsReturnsEmptyWhenAllLabelsUnique", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateLabelsReportsAllDuplicate()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateLabelsReportsAllDuplicate"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTranslationLabel "Hello"
    AppendTranslationLabel "World"
    AppendTranslationLabel "Hello"
    AppendTranslationLabel "World"

    Dim summary As String
    Dim duplicateMessage As String
    Assert.IsTrue Subject.DuplicateLabels(duplicateMessage), "DuplicateLabels should return True when duplicates exist"
    Assert.AreEqual "Duplicate labels detected in column English!" & vbLf & """Hello"" has 2 duplicates"  & vbLf & """World"" has 2 duplicates",  _ 
    duplicateMessage, "DuplicateLabels should list all duplicates for the label column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsReportsFirstDuplicate", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateLabelsHonoursLanguageParameter()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateLabelsHonoursLanguageParameter"
    On Error GoTo Fail

    ResetTranslationsTableRows
    Subject.EnsureLanguages "French"

    AppendTranslationLabel "Alpha"
    AppendTranslationLabel "Beta"
    AppendTranslationLabel "Gamma"

    TranslationsTable.ListColumns("French").DataBodyRange.Cells(1, 1).Value = "Bonjour"
    TranslationsTable.ListColumns("French").DataBodyRange.Cells(2, 1).Value = "Salut"
    TranslationsTable.ListColumns("French").DataBodyRange.Cells(3, 1).Value = "Bonjour"

    Dim summary As String
    Dim frenchSummary As String
    Assert.IsTrue Subject.DuplicateLabels(frenchSummary, "French"), "DuplicateLabels should detect duplicates within the specified language column"
    Assert.AreEqual "Duplicate labels detected in column French!" & vbLf & """Bonjour"" has 2 duplicates", frenchSummary, "DuplicateLabels should evaluate duplicates within the specified language column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsHonoursLanguageParameter", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateLabelsListsAllDuplicateValues()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateLabelsListsAllDuplicateValues"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTranslationLabel "One"
    AppendTranslationLabel "Two"
    AppendTranslationLabel "One"
    AppendTranslationLabel "Three"
    AppendTranslationLabel "Two"

    Dim duplicateMessage As String
    Assert.IsTrue Subject.DuplicateLabels(duplicateMessage), "DuplicateLabels should detect multiple duplicate values"

    Dim expected As String
    expected = "Duplicate labels detected in column English!" & vbLf & _
               """One"" has 2 duplicates" & vbLf & _
               """Two"" has 2 duplicates"

    Assert.AreEqual expected, duplicateMessage, "DuplicateLabels should include each duplicated value in the summary"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsListsAllDuplicateValues", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestMissingLabelsRejectsUnknownLanguage()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestMissingLabelsRejectsUnknownLanguage"

    On Error GoTo ExpectError
        Subject.UpdateFromRegistry RegistrySheet
        Subject.MissingLabels "Spanish"
        Assert.LogFailure "MissingLabels should raise an error when the language does not exist"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "MissingLabels must raise InvalidArgument for unknown languages"
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
        Array("TableName", "rngname", "status", "mode"), _
        Array("table", "RNG_Greetings", "yes", "translate as text"), _
        Array("table", "RNG_Farewell", "no", "translate as text"), _
        Array("table", "RNG_Formula", "yes", "translate as formula")))

    targetSheet.Cells.Clear
    TestHelpers.WriteMatrix targetSheet.Cells(1, 1), matrix

    Dim registryRange As Range
    Set registryRange = targetSheet.Range("A1:D4")

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
    targetSheet.Range("C1").Formula = "IF(A1="""", ""Morning"", ""Evening"")"

    hostWorkbook.Names.Add Name:="RNG_Greetings", RefersTo:=targetSheet.Range("A1:A2")
    hostWorkbook.Names.Add Name:="RNG_Farewell", RefersTo:=targetSheet.Range("B1:B2")
    hostWorkbook.Names.Add Name:="RNG_Formula", RefersTo:=targetSheet.Range("C1")
End Sub

Private Sub SetRegistryStatus(ByVal firstStatus As String, ByVal secondStatus As String, ByVal thirdStatus As String)
    RegistryTable.ListRows(1).Range.Cells(1, 3).Value = firstStatus
    RegistryTable.ListRows(2).Range.Cells(1, 3).Value = secondStatus
    RegistryTable.ListRows(3).Range.Cells(1, 3).Value = thirdStatus
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

Private Sub ResetTranslationsTableRows()
    On Error Resume Next
        Do While TranslationsTable.ListRows.Count > 0
            TranslationsTable.ListRows(TranslationsTable.ListRows.Count).Delete
        Loop
    On Error GoTo 0
End Sub

Private Sub AppendTranslationLabel(ByVal label As String)
    Dim newRow As ListRow
    Set newRow = TranslationsTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = label
End Sub

Private Function HasColumn(ByVal columnName As String) As Boolean
    Dim column As ListColumn
    For Each column In TranslationsTable.ListColumns
        If StrComp(column.Name, columnName, vbTextCompare) = 0 Then
            HasColumn = True
            Exit Function
        End If
    Next column
End Function

Private Function ExpectedTag(ByVal rangeName As String, ByVal sequenceNumber As Long) As String
    ExpectedTag = rangeName & TAG_SEPARATOR & CStr(sequenceNumber)
End Function

Private Function CounterValue() As Long
    Dim store As IHiddenNames

    Set store = HiddenCounterStore()
    If store Is Nothing Then Exit Function

    CounterValue = store.ValueAsLong(COUNTER_NAME, 0)
End Function

Private Function HiddenCounterExists() As Boolean
    Dim store As IHiddenNames

    Set store = HiddenCounterStore()
    If store Is Nothing Then Exit Function

    HiddenCounterExists = store.HasName(COUNTER_NAME)
End Function

Private Function HiddenCounterStore() As IHiddenNames
    On Error Resume Next
        Set HiddenCounterStore = HiddenNames.Create(RegistrySheet)
    On Error GoTo 0
End Function
