Attribute VB_Name = "TestSetupTranslationsTable"
Attribute VB_Description = "Unit tests for the improved translations table manager"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the SetupTranslationsTable class covering caching, registry updates and language management")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed
'@depends SetupTranslationsTable, CustomTest, HiddenNames, BetterArray, ProjectError, Messenger

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private TranslationsSheet As Worksheet
Private RegistrySheet As Worksheet
Private SourceSheet As Worksheet
Private TranslationsTable As ListObject
Private RegistryTable As ListObject
Private Subject As SetupTranslationsTable

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const REGISTRY_SHEET_NAME As String = "Registry"
Private Const SOURCE_SHEET_NAME As String = "SourceData"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const REGISTRY_TABLE_NAME As String = "Tab_Registry"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const TAG_SEPARATOR As String = "__"
'The tag MarkImported writes on every imported row.
Private Const IMPORTED_TAG As String = "__imported__" & TAG_SEPARATOR & "0"
Private Const LANGUAGES_NAME_ID As String = "__SetupTranslationsLanguages__"
Private Const LARGE_RANGE_NAME As String = "RNG_Large"
Private Const MARKER_BELOW_TABLE As String = "Below the table"

'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail

    BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupTranslationsTable"
    Exit Sub

Fail:
    If Not Assert Is Nothing Then
        CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
    End If
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = NewWorkbook
    Set TranslationsSheet = EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)
    Set RegistrySheet = EnsureWorksheet(REGISTRY_SHEET_NAME, FixtureWorkbook)
    Set SourceSheet = EnsureWorksheet(SOURCE_SHEET_NAME, FixtureWorkbook)

    Set TranslationsTable = BuildTranslationsTable(TranslationsSheet)
    Set RegistryTable = BuildRegistryTable(RegistrySheet)
    RegisterSourceRanges SourceSheet, FixtureWorkbook

    Set Subject = SetupTranslationsTable.Create(TranslationsTable)
    Subject.SetDisplayPrompts False
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    'A test that arms the messenger and then dies would leave every later box
    'swallowed. Reset empties the record and turns the boxes back on.
    Messenger.Reset

    On Error Resume Next
        If Not TranslationsSheet Is Nothing Then TranslationsSheet.Unprotect
        DeleteWorkbook FixtureWorkbook
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
        Dim invalid As SetupTranslationsTable
        Set invalid = SetupTranslationsTable.Create(Nothing)
        Assert.LogFailure "Create should reject a missing listobject"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Create must raise InvalidArgument when the listobject is missing"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestLanguagesNameIdAnswersTheHiddenNameKey()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLanguagesNameIdAnswersTheHiddenNameKey"
    On Error GoTo Fail

    'The designer reads the language list of a loaded setup through this
    'key, so the answer is pinned, and it comes off the predeclared instance
    Assert.AreEqual LANGUAGES_NAME_ID, SetupTranslationsTable.LanguagesNameId, _
                    "LanguagesNameId should answer the persisted language list key"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLanguagesNameIdAnswersTheHiddenNameKey", Err.Number, Err.Description
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
Public Sub TestRemoveLanguageDeletesTheColumn()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRemoveLanguageDeletesTheColumn"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German"
    Subject.RemoveLanguage "French"

    Assert.AreEqual CLng(2), TranslationsTable.ListColumns.Count, "One language column should be gone"
    Assert.IsFalse HasColumn("French"), "The removed language column should be gone"
    Assert.IsTrue HasColumn("German"), "The other language column should remain"
    Assert.IsTrue HasColumn("English"), "The default language column should remain"

    Dim languages As BetterArray
    Set languages = Subject.Languages
    Assert.AreEqual CLng(1), languages.Length, "The stored language list should follow the delete"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveLanguageDeletesTheColumn", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestRemoveLanguageRefusesDefaultAndUnknown()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRemoveLanguageRefusesDefaultAndUnknown"

    Dim refusedDefault As Boolean
    Dim refusedUnknown As Boolean

    Subject.EnsureLanguages "French"

    On Error Resume Next
        Subject.RemoveLanguage "English"
        refusedDefault = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
        Subject.RemoveLanguage "Klingon"
        refusedUnknown = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue refusedDefault, "The default language column must stay"
    Assert.IsTrue refusedUnknown, "An unknown language should be refused"
    Assert.AreEqual CLng(2), TranslationsTable.ListColumns.Count, "No column should be deleted by a refusal"
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestRemoveLanguageDeletesSeveralColumns()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRemoveLanguageDeletesSeveralColumns"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German;Italian"
    Subject.RemoveLanguage "French; Italian"

    Assert.AreEqual CLng(2), TranslationsTable.ListColumns.Count, "Two language columns should be gone"
    Assert.IsFalse HasColumn("French"), "The first named language column should be gone"
    Assert.IsFalse HasColumn("Italian"), "The second named language column should be gone"
    Assert.IsTrue HasColumn("German"), "The language left out of the list should remain"
    Assert.IsTrue HasColumn("English"), "The default language column should remain"

    Dim languages As BetterArray
    Set languages = Subject.Languages
    Assert.AreEqual CLng(1), languages.Length, "The stored language list should follow the deletes"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveLanguageDeletesSeveralColumns", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestRemoveLanguageKeepsEveryColumnWhenOneNameIsRefused()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRemoveLanguageKeepsEveryColumnWhenOneNameIsRefused"

    Dim refused As Boolean

    Subject.EnsureLanguages "French;German"

    On Error Resume Next
        Subject.RemoveLanguage "French;Klingon"
        refused = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue refused, "An unknown name in the list should be refused"
    Assert.AreEqual CLng(3), TranslationsTable.ListColumns.Count, "No column goes when one name of the list is refused"
    Assert.IsTrue HasColumn("French"), "The known language of a refused list should remain"
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEnsureLanguagesPersistsHiddenName()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEnsureLanguagesPersistsHiddenName"
    On Error GoTo Fail

    Subject.EnsureLanguages "French;German"

    Dim store As HiddenNames
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
    Dim exportedStore As HiddenNames
    Dim expectedLanguages As String

    On Error GoTo Fail

    Subject.EnsureLanguages "French"
    TranslationsTable.DataBodyRange.Cells(1, 1).Value = "Hello"

    expectedLanguages = HiddenNames.Create(TranslationsSheet).ValueAsString(LANGUAGES_NAME_ID)

    Set exportBook = NewWorkbook
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
    singleMatrix = RowsToMatrix(Array( _
                                    Array("tabname", "rngname", "status", "mode"), _
                                    Array("table", "RNG_Solo", "yes", "translate as text")))
    WriteMatrix RegistrySheet.Range("F1"), singleMatrix

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

'@description
'The closing summary of the tag update goes through Messenger.Show. Arrange:
'updates the table from the registry, turns the prompts back on and arms the
'messenger. Act: calls NumberOfMissing. Assert: the summary comes back the same
'and the messenger holds the text of the box.
'
'The messenger is armed BEFORE the prompts go on. A disarmed messenger opens a
'real box, and a box nobody clicks holds the whole run.
'@TestMethod("SetupTranslationsTable")
Public Sub TestNumberOfMissingSummarySpeaksThroughTheMessenger()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestNumberOfMissingSummarySpeaksThroughTheMessenger"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"
    Messenger.Arm FixtureWorkbook
    Subject.SetDisplayPrompts True

    Dim summary As String
    summary = Subject.NumberOfMissing

    Subject.SetDisplayPrompts False
    Messenger.Disarm

    Assert.AreEqual "Translation Updated!" & vbLf & "6 labels are missing for column French.", summary, "The summary the function answers should not change when the messenger is armed"
    Assert.IsTrue Messenger.HasMessages, "The armed messenger should hold the summary box"
    Assert.IsTrue InStr(1, Messenger.Messages, "Translation Updated!", vbTextCompare) > 0, "The recorded line should carry the text of the summary"
    Assert.IsTrue InStr(1, Messenger.Messages, "Done!", vbTextCompare) > 0, "The recorded line should carry the title of the summary box"
    Exit Sub

Fail:
    Subject.SetDisplayPrompts False
    Messenger.Disarm
    CustomTestLogFailure Assert, "TestNumberOfMissingSummarySpeaksThroughTheMessenger", Err.Number, Err.Description
End Sub

'@description
'A summary with the prompts off records nothing. Arrange: updates the table from
'the registry and arms the messenger, leaving the prompts off the way
'TestInitialize sets them. Act: calls NumberOfMissing. Assert: the messenger
'record is empty.
'
'The displayPrompts guard sits in front of the Show call and it still decides
'whether there is a box at all.
'@TestMethod("SetupTranslationsTable")
Public Sub TestNumberOfMissingWithPromptsOffRecordsNothing()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestNumberOfMissingWithPromptsOffRecordsNothing"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"
    Messenger.Arm FixtureWorkbook

    Subject.NumberOfMissing

    Messenger.Disarm

    Assert.IsFalse Messenger.HasMessages, "A summary raised with the prompts off should record nothing"
    Exit Sub

Fail:
    Messenger.Disarm
    CustomTestLogFailure Assert, "TestNumberOfMissingWithPromptsOffRecordsNothing", Err.Number, Err.Description
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

    Dim duplicateMessage As String
    Assert.IsTrue Subject.DuplicateLabels(duplicateMessage), "DuplicateLabels should return True when duplicates exist"
    Assert.AreEqual "Duplicate labels detected in column English!" & vbLf & """Hello"" has 2 duplicates" & vbLf & """World"" has 2 duplicates", _
                    duplicateMessage, "DuplicateLabels should list all duplicates for the label column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsReportsAllDuplicate", Err.Number, Err.Description
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

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryKeepsEldestDuplicateRow()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryKeepsEldestDuplicateRow"
    On Error GoTo Fail

    'First update creates labels with sequence 1
    Subject.UpdateFromRegistry RegistrySheet

    'Manually add a duplicate "Hello" with an older tag (sequence 0)
    AppendTaggedLabel "Hello", "EXTRA" & TAG_SEPARATOR & "0"

    'EXTRA names no registry range, so the row is unseen and would go; a
    'review answered no (the preset) keeps it for the dedup to judge.
    Subject.RequestUnseenReview

    'Second update with all statuses "yes" triggers dedup
    SetRegistryStatus "yes", "yes", "yes"
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), CountLabelIgnoringCase("Hello"), "Dedup should leave exactly one Hello row"
    Assert.AreEqual "EXTRA" & TAG_SEPARATOR & "0", TagForLabel("Hello"), _
                    "Dedup should keep the row with the oldest (lowest sequence) tag"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryKeepsEldestDuplicateRow", Err.Number, Err.Description
End Sub


'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryRemovesUnseenLabelsSilently()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryRemovesUnseenLabelsSilently"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet

    'An imported label, a row left by a range the registry has dropped, and
    'a hand-written row. No review was requested: the first two go, the
    'hand-written one is the user's and stays.
    AppendTaggedLabel "Imported orphan", IMPORTED_TAG
    AppendTaggedLabel "Retired chunk", "RNG_Retired" & TAG_SEPARATOR & "1"
    AppendTaggedLabel "Handmade", vbNullString
    SetRegistryStatus "no", "no", "no"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(2), Subject.UnseenLabels.Length, "Both rows no registry range produces are reported as unseen"
    Assert.IsFalse LabelExists("Imported orphan"), "An unseen label goes without a question between imports"
    Assert.IsFalse LabelExists("Retired chunk"), "A label of a dropped range goes without a question between imports"
    Assert.IsTrue LabelExists("Handmade"), "A hand-written row is never unseen"
    Assert.AreEqual CLng(7), TranslationsTable.ListRows.Count, "The rows of the registry ranges and the hand-written row all stay"
    Assert.AreEqual ExpectedTag("RNG_Farewell", 2), TagForLabel("Farewell"), _
                    "A cycle with unclaimed rows reads every range, whatever its status"
    Assert.IsTrue InStr(1, Subject.NumberOfMissing(), "2 labels came from no range", vbTextCompare) > 0, _
                  "The summary counts the removed labels"

    'A third update finds nothing left.
    Subject.UpdateFromRegistry RegistrySheet
    Assert.AreEqual CLng(0), Subject.UnseenLabels.Length, "Nothing is unseen once the orphans are gone"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryRemovesUnseenLabelsSilently", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryAsksAboutUnseenLabelsAfterReviewRequest()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryAsksAboutUnseenLabelsAfterReviewRequest"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    AppendTaggedLabel "Imported orphan", IMPORTED_TAG
    AppendTaggedLabel "Retired chunk", "RNG_Retired" & TAG_SEPARATOR & "1"
    AppendTaggedLabel "Handmade", vbNullString
    SetRegistryStatus "no", "no", "no"

    'Reset tags asks the next update to review; prompts are off, so the
    'preset answers, and it keeps the labels by default. The review offers
    'the hand-written row as well.
    Subject.RequestUnseenReview
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(3), Subject.UnseenLabels.Length, "The reviewed labels are reported, the hand-written one included"
    Assert.IsTrue LabelExists("Imported orphan"), "A reviewed label is kept when the answer is no"
    Assert.IsTrue LabelExists("Retired chunk"), "A reviewed label of a dropped range is kept when the answer is no"
    Assert.IsTrue LabelExists("Handmade"), "A reviewed hand-written label is kept when the answer is no"
    Assert.IsTrue InStr(1, Subject.NumberOfMissing(), "3 labels come from no range", vbTextCompare) > 0, _
                  "The summary counts the kept labels"

    'The request is spent: the next update is silent again, removes the
    'imported and retired rows, and leaves the hand-written one alone.
    Subject.UpdateFromRegistry RegistrySheet
    Assert.IsFalse LabelExists("Imported orphan"), "The update after the review removes the unseen labels again"
    Assert.IsFalse LabelExists("Retired chunk"), "The update after the review removes every unseen label"
    Assert.IsTrue LabelExists("Handmade"), "A silent update never removes a hand-written row"

    'A review answered yes removes them at once, the hand-written row too.
    AppendTaggedLabel "Second orphan", IMPORTED_TAG
    Subject.SetRemoveUnseenLabels True
    Subject.RequestUnseenReview
    Subject.UpdateFromRegistry RegistrySheet
    Assert.IsFalse LabelExists("Second orphan"), "A reviewed label goes when the answer is yes"
    Assert.IsFalse LabelExists("Handmade"), "A reviewed hand-written label goes when the answer is yes"
    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "The rows of the registry ranges all stay"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateFromRegistryAsksAboutUnseenLabelsAfterReviewRequest", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestMarkImportedClearsTagsAndRequestsReview()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestMarkImportedClearsTagsAndRequestsReview"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    AppendTaggedLabel "Imported orphan", vbNullString
    SetRegistryStatus "no", "no", "no"

    Subject.MarkImported

    Assert.AreEqual IMPORTED_TAG, TagForLabel("Hello"), "An import marks every row of the helper column as imported"
    Assert.AreEqual IMPORTED_TAG, TagForLabel("Imported orphan"), "A row with no tag is marked as imported too"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual ExpectedTag("RNG_Greetings", 2), TagForLabel("Hello"), "The update after an import tags every row again"
    Assert.AreEqual ExpectedTag("RNG_Farewell", 2), TagForLabel("Farewell"), "The update after an import reads every range"
    Assert.AreEqual CLng(1), Subject.UnseenLabels.Length, "The update after an import reports the unseen labels"
    Assert.IsTrue LabelExists("Imported orphan"), "The update after an import asks before removing, and the preset keeps"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMarkImportedClearsTagsAndRequestsReview", Err.Number, Err.Description
End Sub


'@section Correctness guards
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateKeepsUntaggedRows()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateKeepsUntaggedRows"
    On Error GoTo Fail

    Subject.EnsureLanguages "French"
    ResetTranslationsTableRows
    AppendTaggedLabel "Handmade one", vbNullString
    AppendTaggedLabel "Handmade two", vbNullString
    AppendTaggedLabel "Handmade three", vbNullString
    TranslationsTable.ListColumns("French").DataBodyRange.Cells(1, 1).Value = "Fait main"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Handmade one"), "A row that carries no tag must survive the first update"
    Assert.IsTrue LabelExists("Handmade two"), "Every untagged row must survive the first update"
    Assert.IsTrue LabelExists("Handmade three"), "Every untagged row must survive the first update"
    Assert.AreEqual "Fait main", TranslationForLabel("Handmade one", "French"), _
                    "The translation typed against an untagged row must survive the update"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateKeepsUntaggedRows", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateKeepsLabelsWhenRangeIsMissing()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateKeepsLabelsWhenRangeIsMissing"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    DropName "RNG_Farewell"
    SetRegistryStatus "yes", "yes", "yes"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Farewell"), "A range that fails to resolve must keep the labels tagged to it"
    Assert.IsTrue LabelExists("See you"), "Every label of an unresolved range must stay in the table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateKeepsLabelsWhenRangeIsMissing", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateReportsUnresolvedRange()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateReportsUnresolvedRange"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    DropName "RNG_Farewell"
    SetRegistryStatus "yes", "yes", "yes"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), Subject.UnresolvedRanges.Length, "One registry range failed to resolve"
    Assert.IsTrue InStr(1, Subject.NumberOfMissing, "RNG_Farewell", vbTextCompare) > 0, _
                  "The summary must name the range that could not be found"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateReportsUnresolvedRange", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateSurvivesErrorValueInSource()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateSurvivesErrorValueInSource"
    On Error GoTo Fail

    SourceSheet.Range("A2").Formula = "=NA()"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Hello"), "A cell holding an error value must not stop the other cells importing"
    Assert.IsTrue LabelExists("Farewell"), "The update must run to the end when a source cell holds an error"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateSurvivesErrorValueInSource", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateRejectsRegistryWithTooFewColumns()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateRejectsRegistryWithTooFewColumns"

    Dim shortMatrix As Variant
    Dim shortTable As ListObject

    shortMatrix = RowsToMatrix(Array( _
                                   Array("colname", "rngname"), _
                                   Array("table", "RNG_Greetings")))
    WriteMatrix RegistrySheet.Range("F1"), shortMatrix

    Set shortTable = RegistrySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=RegistrySheet.Range("F1:G2"), XlListObjectHasHeaders:=xlYes)
    shortTable.Name = "Tab_RegistryShort"

    On Error GoTo ExpectError
        Subject.UpdateFromRegistry RegistrySheet
        Assert.LogFailure "A registry table with two columns should be reported"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), Err.Number, _
                    "A registry table with too few columns must raise ErrorUnexpectedState"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateAcceptsMixedCaseWatchMode()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateAcceptsMixedCaseWatchMode"
    On Error GoTo Fail

    RegistryTable.ListRows(1).Range.Cells(1, 4).Value = "Watch For Update"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsFalse LabelExists("Hello"), "A watched range brings no label of its own"
    Assert.IsTrue LabelExists("Farewell"), "A mixed case watch mode must not stop the rest of the registry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateAcceptsMixedCaseWatchMode", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateAcceptsBlankMode()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateAcceptsBlankMode"
    On Error GoTo Fail

    RegistryTable.ListRows(1).Range.Cells(1, 4).ClearContents

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsFalse LabelExists("Hello"), "A blank mode brings no label"
    Assert.IsTrue LabelExists("Farewell"), "A blank mode must not stop the rest of the registry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateAcceptsBlankMode", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestSwitchDefaultLanguageMovesDataAndHeader()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestSwitchDefaultLanguageMovesDataAndHeader"
    On Error GoTo Fail

    Subject.EnsureLanguages "French"
    ResetTranslationsTableRows
    AppendTaggedLabel "Alpha", vbNullString
    AppendTaggedLabel "Beta", vbNullString
    AppendTaggedLabel "Gamma", vbNullString
    AppendTaggedLabel "Delta", vbNullString
    AppendTaggedLabel "Epsilon", vbNullString

    Dim rowIndex As Long
    For rowIndex = 1 To 5
        TranslationsTable.ListColumns("French").DataBodyRange.Cells(rowIndex, 1).Value = "Fr " & PadNumber(rowIndex)
    Next rowIndex

    Subject.SwitchDefaultLanguage "French"

    Assert.AreEqual "French", TranslationsTable.ListColumns(1).Name, "The promoted language must become the first column"
    Assert.AreEqual "English", TranslationsTable.ListColumns(2).Name, "The old default must take the second column"
    Assert.AreEqual "Fr 001", CStr(TranslationsTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value), _
                    "The French values must move with their header"
    Assert.AreEqual "Alpha", CStr(TranslationsTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value), _
                    "The English values must move with their header"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSwitchDefaultLanguageMovesDataAndHeader", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title Two spellings of one word are two labels, so neither reports the other.
'@details Fails against the old ComputeDuplicateSummary, which keyed both
'         spellings to the same slot at SetupTranslationsTable:1662 and told the
'         user to go and fix a duplicate that was never there.
Public Sub TestDuplicateLabelsKeepsCasesApart()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateLabelsKeepsCasesApart"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTaggedLabel "Hello", vbNullString
    AppendTaggedLabel "hello", vbNullString

    Dim duplicateMessage As String
    Assert.IsFalse Subject.DuplicateLabels(duplicateMessage), "Two labels that differ only by case are two labels"
    Assert.AreEqual vbNullString, duplicateMessage, "Nothing is reported when no spelling repeats"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateLabelsKeepsCasesApart", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title Dedup leaves both spellings standing.
'@details Fails against the old DeduplicateLabels, which matched the two
'         spellings at SetupTranslationsTable:1343 and deleted one of the user's
'         rows outright.
Public Sub TestDedupKeepsBothCasesOfALabel()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDedupKeepsBothCasesOfALabel"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    AppendTaggedLabel "hello", "EXTRA" & TAG_SEPARATOR & "5"

    SetRegistryStatus "yes", "yes", "yes"
    Subject.UpdateFromRegistry RegistrySheet

    'Hello, Good bye, Morning, Evening, hello, Farewell and See you.
    Assert.AreEqual CLng(2), CountLabelIgnoringCase("hello"), "Dedup must leave Hello and hello both standing"
    Assert.AreEqual CLng(1), CountLabelMatchingCase("Hello"), "The label the registry wrote keeps its own row"
    Assert.AreEqual CLng(1), CountLabelMatchingCase("hello"), "The label the user typed keeps its own row"
    Assert.AreEqual CLng(7), TranslationsTable.ListRows.Count, "The table keeps all seven labels after dedup"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDedupKeepsBothCasesOfALabel", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title A formula chunk spelled differently gets its own row.
'@details Fails against the old FindLabelRow, which answered the existing row at
'         SetupTranslationsTable:1074, so AddChunk never appended the spelling
'         the formula actually asked for.
Public Sub TestAddChunkCreatesASecondRowForADifferentCase()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestAddChunkCreatesASecondRowForADifferentCase"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTaggedLabel "MORNING", vbNullString

    SetRegistryStatus "no", "no", "yes"
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(2), CountLabelIgnoringCase("morning"), "The formula chunk must add a row of its own"
    Assert.AreEqual CLng(1), CountLabelMatchingCase("MORNING"), "The row already there is left alone"
    Assert.AreEqual CLng(1), CountLabelMatchingCase("Morning"), "The chunk the formula carries gets its own row"
    Assert.IsTrue LabelExists("Evening"), "The rest of the formula is still processed"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddChunkCreatesASecondRowForADifferentCase", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title A watched text cell matches only the row spelled the same way.
'@details Fails against the old FindLabelRow at SetupTranslationsTable:1074, the
'         text path this time: the cell reading hello was answered by the Hello
'         row and no hello row was ever written.
Public Sub TestLabelRowLookupIsCaseSensitive()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLabelRowLookupIsCaseSensitive"
    On Error GoTo Fail

    SourceSheet.Range("A1").Value = "hello"

    ResetTranslationsTableRows
    AppendTaggedLabel "Hello", vbNullString

    SetRegistryStatus "yes", "no", "no"
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), CountLabelMatchingCase("Hello"), "The row already there is left alone"
    Assert.AreEqual CLng(1), CountLabelMatchingCase("hello"), "The watched cell gets the row it is spelled for"
    Assert.IsTrue LabelExists("Good bye"), "The rest of the watched range is still processed"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLabelRowLookupIsCaseSensitive", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title A label that is a bare number still keys and still counts.
'@details The key carries a leading character because a bare number is a legal
'         label and an illegal Collection key. Whatever spells the case out has
'         to leave that prefix in place.
Public Sub TestNumericLabelStillWorks()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestNumericLabelStillWorks"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTaggedLabel "123", vbNullString
    AppendTaggedLabel "123", vbNullString
    AppendTaggedLabel "456", vbNullString

    Dim duplicateMessage As String
    Assert.IsTrue Subject.DuplicateLabels(duplicateMessage), "A numeric label that repeats is still a duplicate"
    Assert.IsTrue InStr(1, duplicateMessage, "123", vbBinaryCompare) > 0, "The repeated number must be named"
    Assert.IsTrue InStr(1, duplicateMessage, "456", vbBinaryCompare) = 0, "The number that appears once is not a duplicate"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestNumericLabelStillWorks", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestParseTagAcceptsRangeNameWithSeparator()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestParseTagAcceptsRangeNameWithSeparator"
    On Error GoTo Fail

    SourceSheet.Range("E1").Value = "Sep label"
    DropName "RNG_a__b"
    FixtureWorkbook.Names.Add Name:="RNG_a__b", RefersTo:=SourceSheet.Range("E1")

    ResetTranslationsTableRows
    AppendTaggedLabel "Stale sep", "RNG_a__b" & TAG_SEPARATOR & "3"

    RegistryTable.ListRows(1).Range.Cells(1, 2).Value = "RNG_a__b"
    SetRegistryStatus "yes", "no", "no"
    SetCounterValue 3

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Sep label"), "The range whose name carries the separator must be processed"
    Assert.IsFalse LabelExists("Stale sep"), "A tag whose range name carries the separator must still parse as stale"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestParseTagAcceptsRangeNameWithSeparator", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestCounterHoldsAfterFailedUpdate()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestCounterHoldsAfterFailedUpdate"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    RegistryTable.ListRows(1).Range.Cells(1, 4).Value = "unsupported"
    SetRegistryStatus "yes", "no", "no"

    RunFailingUpdate

    Assert.AreEqual CLng(1), CounterValue(), "A failed update must leave the counter where it was"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCounterHoldsAfterFailedUpdate", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateDetectsShiftedHelperColumn()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateDetectsShiftedHelperColumn"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    TranslationsSheet.Columns(2).Insert Shift:=xlToRight
    SetRegistryStatus "yes", "yes", "yes"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual "English", TranslationsTable.ListColumns(1).Name, _
                    "The label column keeps its name when a column is inserted beside the helper"
    Assert.IsTrue LabelExists("Hello"), "Labels survive a helper column that moved away from the table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateDetectsShiftedHelperColumn", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateRestoresScreenAfterFailure()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateRestoresScreenAfterFailure"
    On Error GoTo Fail

    Dim screenBefore As Boolean
    Dim reported As Long

    screenBefore = ScreenStateSnapshot()
    RegistryTable.ListRows(1).Range.Cells(1, 4).Value = "unsupported"

    reported = RunFailingUpdate()

    AssertScreenRestored screenBefore, "UpdateFromRegistry"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), reported, _
                    "The error the caller sees must be the real one, not one raised by the cleanup path"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateRestoresScreenAfterFailure", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateEndsTagIntegrationAfterFailure()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateEndsTagIntegrationAfterFailure"
    On Error GoTo Fail

    Dim columnsBefore As Long

    columnsBefore = TranslationsTable.ListColumns.Count
    RegistryTable.ListRows(1).Range.Cells(1, 4).Value = "unsupported"

    RunFailingUpdate

    Assert.AreEqual columnsBefore, TranslationsTable.ListColumns.Count, _
                    "The helper column must be back outside the table after a failed update"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateEndsTagIntegrationAfterFailure", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEnsureLanguagesKeepsTypedNamesWhenAddFails()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEnsureLanguagesKeepsTypedNamesWhenAddFails"
    On Error GoTo Fail

    Dim typedCell As Range

    Set typedCell = TranslationsTable.ListColumns(TranslationsTable.ListColumns.Count).Range.Offset(, 1).Cells(1, 1)
    typedCell.Value = "German"
    TranslationsSheet.Protect

    RunFailingEnsureLanguages

    TranslationsSheet.Unprotect

    Assert.AreEqual "German", CStr(typedCell.Value), _
                    "A language name typed by the user must survive a column add that fails"
    Exit Sub

Fail:
    On Error Resume Next
        TranslationsSheet.Unprotect
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestEnsureLanguagesKeepsTypedNamesWhenAddFails", Err.Number, Err.Description
End Sub


'@section Speed guards
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestLargeTableKeepsEveryLabelAndTag()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLargeTableKeepsEveryLabelAndTag"
    On Error GoTo Fail

    Dim started As Single
    Dim sampleIndex As Long
    Dim sampledLabel As String

    BuildLargeSource 500

    started = Timer
    Subject.UpdateFromRegistry RegistrySheet
    Assert.LogSuccesses "TestLargeTableKeepsEveryLabelAndTag: 500 labels in " & Format$(Timer - started, "0.000") & "s"

    Assert.AreEqual CLng(500), TranslationsTable.ListRows.Count, "Every source label must reach the table"

    For sampleIndex = 1 To 10
        sampledLabel = LargeLabel(sampleIndex * 37)
        Assert.AreEqual ExpectedTag(LARGE_RANGE_NAME, 1), TagForLabel(sampledLabel), _
                        "Sampled label " & sampledLabel & " must carry its range tag"
    Next sampleIndex
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLargeTableKeepsEveryLabelAndTag", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestLargeTableDeduplicatesCorrectly()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLargeTableDeduplicatesCorrectly"
    On Error GoTo Fail

    Dim started As Single
    Dim plantIndex As Long

    BuildLargeSource 500
    Subject.UpdateFromRegistry RegistrySheet

    'Plant 50 duplicates carrying a HIGHER sequence, so the original row wins.
    For plantIndex = 1 To 50
        AppendTaggedLabel LargeLabel(plantIndex), "EXTRA" & TAG_SEPARATOR & "9"
    Next plantIndex

    started = Timer
    Subject.UpdateFromRegistry RegistrySheet
    Assert.LogSuccesses "TestLargeTableDeduplicatesCorrectly: 550 rows deduped in " & Format$(Timer - started, "0.000") & "s"

    Assert.AreEqual CLng(500), TranslationsTable.ListRows.Count, "Dedup must leave one row per label"
    Assert.AreEqual ExpectedTag(LARGE_RANGE_NAME, 2), TagForLabel(LargeLabel(1)), _
                    "The surviving row of each pair is the one carrying the lower sequence"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLargeTableDeduplicatesCorrectly", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestRemoveObsoleteDeletesOnlyStaleRows()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRemoveObsoleteDeletesOnlyStaleRows"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    SourceSheet.Range("B1:B2").ClearContents
    SetRegistryStatus "yes", "yes", "yes"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(4), TranslationsTable.ListRows.Count, "Only the retired range loses its rows"
    Assert.IsFalse LabelExists("Farewell"), "A label whose range no longer produces it must go"
    Assert.IsFalse LabelExists("See you"), "Every label of the retired range must go"
    Assert.IsTrue LabelExists("Hello"), "A label from another range must stay"
    Assert.IsTrue LabelExists("Morning"), "A label from the formula range must stay"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveObsoleteDeletesOnlyStaleRows", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateSummaryMatchesPerLanguage()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateSummaryMatchesPerLanguage"
    On Error GoTo Fail

    ResetTranslationsTableRows
    Subject.EnsureLanguages "French"
    AppendTaggedLabel "Alpha", vbNullString
    AppendTaggedLabel "Beta", vbNullString
    AppendTaggedLabel "Gamma", vbNullString

    TranslationsTable.ListColumns("French").DataBodyRange.Cells(1, 1).Value = "Bonjour"
    TranslationsTable.ListColumns("French").DataBodyRange.Cells(2, 1).Value = "Salut"
    TranslationsTable.ListColumns("French").DataBodyRange.Cells(3, 1).Value = "Bonjour"

    Dim englishSummary As String
    Dim frenchSummary As String

    Assert.IsFalse Subject.DuplicateLabels(englishSummary), "The label column holds no duplicate"
    Assert.IsTrue Subject.DuplicateLabels(frenchSummary, "French"), "The French column holds one duplicate"
    Assert.IsTrue InStr(1, frenchSummary, "French", vbTextCompare) > 0, "The summary must name the column it read"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateSummaryMatchesPerLanguage", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestSwitchDefaultLanguageOnManyRows()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestSwitchDefaultLanguageOnManyRows"
    On Error GoTo Fail

    Dim started As Single
    Dim frenchBefore As Variant
    Dim frenchAfter As Variant

    Subject.EnsureLanguages "French"
    FillTranslationsRows 200
    frenchBefore = SnapshotColumnValues("French")

    started = Timer
    Subject.SwitchDefaultLanguage "French"
    Assert.LogSuccesses "TestSwitchDefaultLanguageOnManyRows: 200 rows swapped in " & Format$(Timer - started, "0.000") & "s"

    frenchAfter = SnapshotColumnValues("French")
    Assert.AreEqual CStr(frenchBefore(200, 1)), CStr(frenchAfter(200, 1)), _
                    "Every French value must sit on the same row after the swap"
    Assert.AreEqual "Fr 001", CStr(TranslationsTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value), _
                    "The first row must carry its French value after the swap"
    Assert.AreEqual "Fr 200", CStr(TranslationsTable.ListColumns(1).DataBodyRange.Cells(200, 1).Value), _
                    "The last row must carry its French value after the swap"
    Assert.AreEqual "Row 200", CStr(TranslationsTable.ListColumns(2).DataBodyRange.Cells(200, 1).Value), _
                    "The last row must carry its English value after the swap"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSwitchDefaultLanguageOnManyRows", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestFormatConditionsDoNotAccumulate()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestFormatConditionsDoNotAccumulate"
    On Error GoTo Fail

    Dim countBefore As Long
    Dim countAfter As Long
    Dim runIndex As Long

    Subject.UpdateFromRegistry RegistrySheet, "French"
    countBefore = TranslationsTable.Range.FormatConditions.Count

    For runIndex = 1 To 5
        SetRegistryStatus "yes", "yes", "yes"
        Subject.UpdateFromRegistry RegistrySheet
    Next runIndex

    countAfter = TranslationsTable.Range.FormatConditions.Count
    Assert.LogSuccesses "TestFormatConditionsDoNotAccumulate: before=" & CStr(countBefore) & " after=" & CStr(countAfter)

    Assert.AreEqual countBefore, countAfter, "Repeated updates must not pile up conditional format rules"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestFormatConditionsDoNotAccumulate", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title Every duplicate rule stays inside the one language column it was given.
'@details A duplicate is a duplicate WITHIN a language. The same label standing in
'         English and in an untranslated French cell is one label seen twice, not a
'         duplicate, and highlighting both paints the row.
'
'         ApplyFormatting adds its rules one column at a time, so the intent was
'         always per column. What this pins is what the rules are left applying to
'         afterwards: two adjacent columns carrying the same rule, the same fill and
'         the same label are what Excel coalesces into one rule over both columns,
'         and a duplicate-values rule over two columns compares across them.
Public Sub TestDuplicateRulesStayInsideOneColumn()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateRulesStayInsideOneColumn"
    On Error GoTo Fail

    Dim idx As Long
    Dim ruleRange As Range
    Dim widest As Long
    Dim ruleCount As Long

    'Two English labels translated to one French word is a duplicate inside
    'French and no duplicate at all inside English. UpdateFromRegistry
    'deduplicates the label column, so the language column is where a duplicate
    'can stand.
    Subject.UpdateFromRegistry RegistrySheet, "French"
    FillLanguageColumn "French", Array("Salut", "Salut")

    RunUpdateThatAddsARow

    ruleCount = TranslationsTable.Range.FormatConditions.Count
    Assert.IsTrue ruleCount > 0, "The update must leave the duplicate rules on the table"

    For idx = 1 To ruleCount
        Set ruleRange = TranslationsTable.Range.FormatConditions(idx).AppliesTo
        If Not ruleRange Is Nothing Then
            If ruleRange.Columns.Count > widest Then widest = ruleRange.Columns.Count
        End If
    Next idx

    Assert.LogSuccesses "TestDuplicateRulesStayInsideOneColumn: rules=" & CStr(ruleCount) & _
                        " widest=" & CStr(widest)

    Assert.AreEqual CLng(1), CLng(widest), _
                    "No duplicate rule may apply to more than one language column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateRulesStayInsideOneColumn", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
'@sub-title Two language columns never draw a duplicate group in the same colour.
'@details The palette counter used to start at zero on every column, so the first
'         duplicate group of every language drew the same colour. Untranslated
'         columns carry the same spelling, so those same-coloured cells landed on
'         the same rows and the sheet read as though whole rows were highlighted.
'
'         The label column is deduplicated by every update, so it can hold no
'         duplicate group at all. Two further languages are what puts two groups
'         on the sheet at once.
Public Sub TestDuplicateGroupColoursDifferBetweenColumns()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateGroupColoursDifferBetweenColumns"
    On Error GoTo Fail

    Dim frenchColors As BetterArray
    Dim spanishColors As BetterArray
    Dim frenchIndex As Long
    Dim spanishIndex As Long
    Dim sharedCount As Long

    Subject.UpdateFromRegistry RegistrySheet, "French;Spanish"

    'Both languages left untranslated, so both carry the same duplicated spelling
    'on the same rows. That is the sheet the row-wide colour showed up on.
    FillLanguageColumn "French", Array("Hello", "Hello")
    FillLanguageColumn "Spanish", Array("Hello", "Hello")

    RunUpdateThatAddsARow

    Set frenchColors = GroupRuleColors("French")
    Set spanishColors = GroupRuleColors("Spanish")

    Assert.IsTrue frenchColors.Length > 0, "French holds a duplicate group, so it must carry a group rule"
    Assert.IsTrue spanishColors.Length > 0, "Spanish holds a duplicate group, so it must carry a group rule"

    For frenchIndex = frenchColors.LowerBound To frenchColors.UpperBound
        For spanishIndex = spanishColors.LowerBound To spanishColors.UpperBound
            If CLng(frenchColors.Item(frenchIndex)) = CLng(spanishColors.Item(spanishIndex)) Then
                sharedCount = sharedCount + 1
            End If
        Next spanishIndex
    Next frenchIndex

    Assert.LogSuccesses "TestDuplicateGroupColoursDifferBetweenColumns: french=" & _
                        CStr(frenchColors.Length) & " spanish=" & CStr(spanishColors.Length) & _
                        " shared=" & CStr(sharedCount)

    Assert.AreEqual CLng(0), CLng(sharedCount), _
                    "A duplicate group in one language must not draw the colour another language used"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateGroupColoursDifferBetweenColumns", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestFormulaChunksUnchangedAfterRewrite()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestFormulaChunksUnchangedAfterRewrite"
    On Error GoTo Fail

    SourceSheet.Range("C2").Formula = "=CONCATENATE(""alpha"",""beta"")"
    SourceSheet.Range("C3").Formula = "=""ex""&""why"""
    DropName "RNG_Formula"
    FixtureWorkbook.Names.Add Name:="RNG_Formula", RefersTo:=SourceSheet.Range("C1:C3")

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Morning"), "The IF branch text must still be extracted"
    Assert.IsTrue LabelExists("Evening"), "Both IF branches must still be extracted"
    Assert.IsTrue LabelExists("alpha"), "A CONCATENATE argument must still be extracted"
    Assert.IsTrue LabelExists("beta"), "Every CONCATENATE argument must still be extracted"
    Assert.IsTrue LabelExists("ex"), "A joined literal must still be extracted"
    Assert.IsTrue LabelExists("why"), "Every joined literal must still be extracted"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestFormulaChunksUnchangedAfterRewrite", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestWholeColumnNamedRangeIsHandled()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestWholeColumnNamedRangeIsHandled"
    On Error GoTo Fail

    SourceSheet.Range("E1").Value = "Whole one"
    SourceSheet.Range("E2").Value = "Whole two"
    SourceSheet.Range("E3").Value = "Whole three"

    DropName "RNG_Whole"
    FixtureWorkbook.Names.Add Name:="RNG_Whole", RefersTo:=SourceSheet.Range("E:E")
    RegistryTable.ListRows(1).Range.Cells(1, 2).Value = "RNG_Whole"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Whole one"), "A name over a whole column must still import its first value"
    Assert.IsTrue LabelExists("Whole three"), "A name over a whole column must import every value it holds"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestWholeColumnNamedRangeIsHandled", Err.Number, Err.Description
End Sub


'@section Structure guards
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestCreateDoesNotWriteToWorkbook()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestCreateDoesNotWriteToWorkbook"
    On Error GoTo Fail

    Dim helperCell As Range
    Dim other As SetupTranslationsTable

    Set helperCell = TranslationsSheet.Cells(1, 1)
    helperCell.Value = "Untouched"
    helperCell.Font.Color = vbBlack

    Set other = SetupTranslationsTable.Create(TranslationsTable)
    other.SetDisplayPrompts False

    Assert.AreEqual "Untouched", CStr(helperCell.Value), "Create must write no cell"
    Assert.AreEqual CLng(vbBlack), CLng(helperCell.Font.Color), "Create must not repaint the helper header"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateDoesNotWriteToWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestCreateWorksOnProtectedSheet()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestCreateWorksOnProtectedSheet"
    On Error GoTo Fail

    Dim other As SetupTranslationsTable

    TranslationsSheet.Protect
    Set other = SetupTranslationsTable.Create(TranslationsTable)
    TranslationsSheet.Unprotect

    Assert.IsTrue Not other Is Nothing, "Create must succeed against a protected sheet"
    Exit Sub

Fail:
    On Error Resume Next
        TranslationsSheet.Unprotect
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateWorksOnProtectedSheet", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestExportLeavesSourceTableUnchanged()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestExportLeavesSourceTableUnchanged"
    Dim exportBook As Workbook
    Dim rowsBefore As Long
    Dim columnsBefore As Long

    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"
    rowsBefore = TranslationsTable.ListRows.Count
    columnsBefore = TranslationsTable.ListColumns.Count

    Set exportBook = NewWorkbook
    Subject.Export exportBook

    Assert.AreEqual rowsBefore, TranslationsTable.ListRows.Count, "Export must leave the source table row count alone"
    Assert.AreEqual columnsBefore, TranslationsTable.ListColumns.Count, "Export must leave the source table column count alone"

    exportBook.Close SaveChanges:=False
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportLeavesSourceTableUnchanged", Err.Number, Err.Description
    On Error Resume Next
        If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEnsureLanguagesPersistsWhenNothingIsAdded()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEnsureLanguagesPersistsWhenNothingIsAdded"
    On Error GoTo Fail

    Dim store As HiddenNames

    Subject.EnsureLanguages "French"

    Set store = HiddenNames.Create(TranslationsSheet)
    store.SetValue LANGUAGES_NAME_ID, "stale value"

    Subject.EnsureLanguages

    Assert.AreEqual "English;French", store.ValueAsString(LANGUAGES_NAME_ID), _
                    "The hidden language list must be rewritten even when no column is added"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureLanguagesPersistsWhenNothingIsAdded", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestRegistryColumnsResolvedByHeaderName()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestRegistryColumnsResolvedByHeaderName"
    On Error GoTo Fail

    Dim matrix As Variant

    RemoveRegistryTables
    RegistrySheet.Cells.Clear

    matrix = RowsToMatrix(Array( _
                              Array("updated", "headername", "colname", "rngname"), _
                              Array("yes", "translate as text", "table", "RNG_Greetings")))
    WriteMatrix RegistrySheet.Cells(1, 1), matrix

    Set RegistryTable = RegistrySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=RegistrySheet.Range("A1:D2"), XlListObjectHasHeaders:=xlYes)
    RegistryTable.Name = REGISTRY_TABLE_NAME

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Hello"), "The range column must be found by its header name whatever its position"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Hello"), _
                    "The tag must carry the range name read from the rngname column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRegistryColumnsResolvedByHeaderName", Err.Number, Err.Description
End Sub


'@section Edge cases
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateAddsOneRowForRepeatedSourceLabel()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateAddsOneRowForRepeatedSourceLabel"
    On Error GoTo Fail

    SourceSheet.Range("A1").Value = "Twice"
    SourceSheet.Range("A2").Value = "Twice"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), CountLabelIgnoringCase("Twice"), _
                    "A label that appears twice in one source range must land on one row"
    Assert.AreEqual CLng(5), TranslationsTable.ListRows.Count, "The repeated label costs one row"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Twice"), "The one row must carry its range tag"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateAddsOneRowForRepeatedSourceLabel", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateHandlesNumericLabel()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateHandlesNumericLabel"
    On Error GoTo Fail

    SourceSheet.Range("A1").Value = 2024
    SourceSheet.Range("A2").Value = 2024

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), CountLabelIgnoringCase("2024"), _
                    "A label made only of digits must still be looked up as one label"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("2024"), "A numeric label must carry its range tag"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateHandlesNumericLabel", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestLargeTableDeletesMoreThanOneBatch()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestLargeTableDeletesMoreThanOneBatch"
    On Error GoTo Fail

    Dim started As Single

    BuildLargeSource 600
    Subject.UpdateFromRegistry RegistrySheet

    'Shrink the watched range to its first 50 cells, so 550 rows go stale in
    'one call and the delete has to flush more than one batch.
    DropName LARGE_RANGE_NAME
    FixtureWorkbook.Names.Add Name:=LARGE_RANGE_NAME, _
                              RefersTo:=SourceSheet.Range(SourceSheet.Cells(1, 1), SourceSheet.Cells(50, 1))

    started = Timer
    Subject.UpdateFromRegistry RegistrySheet
    Assert.LogSuccesses "TestLargeTableDeletesMoreThanOneBatch: 550 rows deleted in " & Format$(Timer - started, "0.000") & "s"

    Assert.AreEqual CLng(50), TranslationsTable.ListRows.Count, "Only the rows the range no longer produces may go"
    Assert.IsTrue LabelExists(LargeLabel(1)), "The first surviving label must stay"
    Assert.IsTrue LabelExists(LargeLabel(50)), "The last surviving label must stay"
    Assert.IsFalse LabelExists(LargeLabel(600)), "A label past the shrunk range must go"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLargeTableDeletesMoreThanOneBatch", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateOnEmptyTranslationsTable()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateOnEmptyTranslationsTable"
    On Error GoTo Fail

    Dim duplicateMessage As String

    ResetTranslationsTableRows
    Assert.IsFalse Subject.DuplicateLabels(duplicateMessage), "A table holding no data row reports no duplicate"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "An update against an empty table must fill it"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Hello"), "Every row written into an empty table must carry its tag"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateOnEmptyTranslationsTable", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateSkipsRangeOutsideUsedRange()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateSkipsRangeOutsideUsedRange"
    On Error GoTo Fail

    DropName "RNG_Far"
    FixtureWorkbook.Names.Add Name:="RNG_Far", RefersTo:=SourceSheet.Range("Z900:Z910")
    RegistryTable.ListRows(1).Range.Cells(1, 2).Value = "RNG_Far"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsFalse LabelExists("Hello"), "A range that sits outside the used cells brings no label"
    Assert.IsTrue LabelExists("Farewell"), "A range outside the used cells must not stop the rest of the registry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateSkipsRangeOutsideUsedRange", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateKeepsRowWithUnparsableTag()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateKeepsRowWithUnparsableTag"
    On Error GoTo Fail

    ResetTranslationsTableRows
    AppendTaggedLabel "Odd tag row", "junk"
    AppendTaggedLabel "Bad sequence row", "RNG_Greetings" & TAG_SEPARATOR & "abc"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.IsTrue LabelExists("Odd tag row"), "A tag with no separator must leave its row alone"
    Assert.IsTrue LabelExists("Bad sequence row"), "A tag whose sequence is not a number must leave its row alone"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateKeepsRowWithUnparsableTag", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateIgnoresWhitespaceOnlySourceCell()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateIgnoresWhitespaceOnlySourceCell"
    On Error GoTo Fail

    SourceSheet.Range("A2").Value = "   "

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(5), TranslationsTable.ListRows.Count, "A cell holding only spaces brings no label"
    Assert.IsTrue LabelExists("Hello"), "The rest of the range still imports"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateIgnoresWhitespaceOnlySourceCell", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUnresolvedRangesClearOnNextRun()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUnresolvedRangesClearOnNextRun"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet
    DropName "RNG_Farewell"
    SetRegistryStatus "yes", "yes", "yes"
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), Subject.UnresolvedRanges.Length, "The failed range must be reported"

    FixtureWorkbook.Names.Add Name:="RNG_Farewell", RefersTo:=SourceSheet.Range("B1:B2")
    SetRegistryStatus "yes", "yes", "yes"
    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(0), Subject.UnresolvedRanges.Length, "A clean run must clear the report from the run before it"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUnresolvedRangesClearOnNextRun", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateAcceptsRegistrySheetWithNoTables()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateAcceptsRegistrySheetWithNoTables"
    On Error GoTo Fail

    RemoveRegistryTables

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(1), CounterValue(), "A registry sheet holding no table still completes the update"
    Assert.IsFalse LabelExists("Hello"), "A registry sheet holding no table brings no label"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateAcceptsRegistrySheetWithNoTables", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestSwitchDefaultLanguageRejectsUnknownLanguage()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestSwitchDefaultLanguageRejectsUnknownLanguage"

    On Error GoTo ExpectError
        Subject.SwitchDefaultLanguage "Spanish"
        Assert.LogFailure "SwitchDefaultLanguage should reject a language the table does not hold"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, _
                    "An unknown language must raise InvalidArgument"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestSwitchDefaultLanguageIgnoresCurrentDefault()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestSwitchDefaultLanguageIgnoresCurrentDefault"
    On Error GoTo Fail

    Subject.EnsureLanguages "French"
    Subject.SwitchDefaultLanguage "English"

    Assert.AreEqual "English", TranslationsTable.ListColumns(1).Name, "Promoting the current default must change nothing"
    Assert.AreEqual "French", TranslationsTable.ListColumns(2).Name, "The other language must stay where it was"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSwitchDefaultLanguageIgnoresCurrentDefault", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestExportRejectsMissingWorkbook()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestExportRejectsMissingWorkbook"

    On Error GoTo ExpectError
        Subject.Export Nothing
        Assert.LogFailure "Export should reject a missing destination workbook"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), Err.Number, _
                    "A missing destination workbook must raise ObjectNotInitialized"
    Err.Clear
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateFromRegistryRejectsMissingRegistrySheet()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateFromRegistryRejectsMissingRegistrySheet"

    On Error GoTo ExpectError
        Subject.UpdateFromRegistry Nothing
        Assert.LogFailure "UpdateFromRegistry should reject a missing registry sheet"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, _
                    "A missing registry sheet must raise InvalidArgument"
    Err.Clear
End Sub


'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateKeepsCellsBelowTheTable()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateKeepsCellsBelowTheTable"
    On Error GoTo Fail

    Dim markerRow As Long
    Dim labelColumn As Long
    Dim found As Range

    ResetTranslationsTableRows

    labelColumn = TranslationsTable.Range.Column
    markerRow = TranslationsTable.Range.Row + TranslationsTable.Range.Rows.Count
    TranslationsSheet.Cells(markerRow, labelColumn).Value = MARKER_BELOW_TABLE

    Subject.UpdateFromRegistry RegistrySheet

    Set found = TranslationsSheet.Columns(labelColumn).Find(What:=MARKER_BELOW_TABLE, LookIn:=xlValues, LookAt:=xlWhole)

    Assert.IsTrue Not found Is Nothing, "A value sitting under the table must survive an update that adds rows"
    If found Is Nothing Then Exit Sub

    Assert.IsTrue found.Row > LastTableRow(), "The value must still sit outside the table"
    Assert.IsTrue LabelExists("Hello"), "The update still imports its own labels"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateKeepsCellsBelowTheTable", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestUpdateSkipsTheNameIndexTable()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestUpdateSkipsTheNameIndexTable"
    On Error GoTo Fail

    Dim indexMatrix As Variant
    Dim indexTable As ListObject

    'UpdatedValues keeps this table on the same sheet as the registries. It
    'holds three columns and no rngname header, so it is not a registry.
    indexMatrix = RowsToMatrix(Array( _
                                   Array("sheet", "listobject", "registry"), _
                                   Array("Dictionary", "UpLo_Dictionary", "Tab_Registry")))
    WriteMatrix RegistrySheet.Range("F1"), indexMatrix

    Set indexTable = RegistrySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=RegistrySheet.Range("F1:H2"), XlListObjectHasHeaders:=xlYes)
    indexTable.Name = "__UpLo__Names__"

    Subject.UpdateFromRegistry RegistrySheet

    Assert.AreEqual CLng(6), TranslationsTable.ListRows.Count, "The name index table must not stop the registries being read"
    Assert.AreEqual ExpectedTag("RNG_Greetings", 1), TagForLabel("Hello"), "The real registry still tags its labels"
    Assert.AreEqual CLng(0), Subject.UnresolvedRanges.Length, "The name index table must not be read as a registry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUpdateSkipsTheNameIndexTable", Err.Number, Err.Description
End Sub


'@section Tag column rendering
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestTagColumnStaysInvisibleAfterUpdate()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestTagColumnStaysInvisibleAfterUpdate"
    On Error GoTo Fail

    Dim header As Range
    Dim tagData As Range

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Set header = TagHeaderCell()
    Set tagData = TagDataRange()

    Assert.AreEqual "__TagInternal__", CStr(header.Value), "The helper header keeps its marker title"
    Assert.AreEqual CLng(vbWhite), CLng(header.Font.Color), "The helper header text is white"
    Assert.AreEqual CLng(vbWhite), CLng(header.Interior.Color), "The helper header fill is white"
    Assert.AreEqual CLng(vbWhite), CLng(tagData.Font.Color), "Every tag reads white on white"

    AssertTagBordersAreWhite tagData, "after an update that changed rows"

    Assert.LogSuccesses "TestTagColumnStaysInvisibleAfterUpdate: tag column width = " & _
                        Format$(header.EntireColumn.ColumnWidth, "0.0") & _
                        ", table column width = " & Format$(TranslationsTable.Range.Columns(1).ColumnWidth, "0.0")
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTagColumnStaysInvisibleAfterUpdate", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestTagColumnStaysInvisibleWhenNothingChanges()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestTagColumnStaysInvisibleWhenNothingChanges"
    On Error GoTo Fail

    Dim header As Range
    Dim tagData As Range

    'The first update adds rows, so ApplyFormatting runs. The second adds and
    'removes nothing, so ApplyFormatting is skipped and EndTagIntegration is the
    'only thing left painting the tag column.
    Subject.UpdateFromRegistry RegistrySheet
    SetRegistryStatus "no", "no", "no"

    'Repaint the tag column in black so the update has something to undo.
    Set header = TagHeaderCell()
    Set tagData = TagDataRange()
    header.Font.Color = vbBlack
    tagData.Font.Color = vbBlack

    Subject.UpdateFromRegistry RegistrySheet

    Set header = TagHeaderCell()
    Set tagData = TagDataRange()

    Assert.AreEqual CLng(vbWhite), CLng(header.Font.Color), "An update that changes no row still hides the header"
    Assert.AreEqual CLng(vbWhite), CLng(header.Interior.Color), "An update that changes no row still fills the header white"
    Assert.AreEqual CLng(vbWhite), CLng(tagData.Font.Color), "An update that changes no row still hides every tag"

    AssertTagBordersAreWhite tagData, "after an update that changed no row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTagColumnStaysInvisibleWhenNothingChanges", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestTableHeaderRowIsWhite()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestTableHeaderRowIsWhite"
    On Error GoTo Fail

    Subject.UpdateFromRegistry RegistrySheet, "French"

    Assert.AreEqual "TableStyleLight11", CStr(TranslationsTable.TableStyle), "The table keeps the style that paints the header band"
    Assert.AreEqual CLng(vbWhite), CLng(TranslationsTable.HeaderRowRange.Font.Color), _
                    "The header row text is white, which is what reads on the band"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTableHeaderRowIsWhite", Err.Number, Err.Description
End Sub


'@section Duplicate group colouring
'===============================================================================
'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateGroupColorsAreDistinct()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateGroupColorsAreDistinct"
    On Error GoTo Fail

    Dim paletteSize As Long
    Dim idx As Long
    Dim jdx As Long
    Dim collisions As Long
    Dim redHits As Long
    Dim holdsSeveral As Boolean

    paletteSize = Subject.DuplicateColorCount
    holdsSeveral = (paletteSize > 1)

    Assert.IsTrue holdsSeveral, "The palette must hold more than one colour"

    For idx = 0 To paletteSize - 1
        If Subject.DuplicateGroupColor(idx) = CLng(vbRed) Then redHits = redHits + 1
        For jdx = idx + 1 To paletteSize - 1
            If Subject.DuplicateGroupColor(idx) = Subject.DuplicateGroupColor(jdx) Then
                collisions = collisions + 1
            End If
        Next jdx
    Next idx

    Assert.AreEqual CLng(0), collisions, "Every group inside the palette must take a colour of its own"
    Assert.AreEqual CLng(0), redHits, "No group colour may be the red the catch-all rule uses"
    Assert.AreEqual Subject.DuplicateGroupColor(0), Subject.DuplicateGroupColor(paletteSize), _
                    "The palette wraps once it is walked through"
    Assert.AreEqual Subject.DuplicateGroupColor(0), Subject.DuplicateGroupColor(-3), _
                    "A negative index answers the first colour rather than raising"

    Assert.LogSuccesses "TestDuplicateGroupColorsAreDistinct: palette holds " & CStr(paletteSize) & " colours"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateGroupColorsAreDistinct", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestEachDuplicateGroupTakesItsOwnColor()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestEachDuplicateGroupTakesItsOwnColor"
    On Error GoTo Fail

    Dim colors As BetterArray
    Dim firstFill As Long
    Dim secondFill As Long
    Dim fillsDiffer As Boolean

    Subject.UpdateFromRegistry RegistrySheet, "French"

    'Two groups: "Bonjour" three times and "Salut" twice. The rest stand alone.
    FillLanguageColumn "French", Array("Bonjour", "Salut", "Bonjour", "Salut", "Bonjour", "Adieu")
    RunUpdateThatAddsARow

    Set colors = GroupRuleColors("French")

    Assert.AreEqual CLng(2), colors.Length, "Two duplicate groups must produce two group rules"
    If colors.Length < 2 Then Exit Sub

    firstFill = CLng(colors.Item(colors.LowerBound))
    secondFill = CLng(colors.Item(colors.LowerBound + 1))
    fillsDiffer = (firstFill <> secondFill)

    Assert.IsTrue fillsDiffer, "The two groups must not share a fill"
    Assert.IsTrue IsPaletteColor(firstFill), "The first group fill comes from the palette"
    Assert.IsTrue IsPaletteColor(secondFill), "The second group fill comes from the palette"

    Assert.AreEqual CLng(vbRed), CLng(LastRuleColor("French")), _
                    "The red catch-all stays on, added last so it holds the lowest priority"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEachDuplicateGroupTakesItsOwnColor", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestCaseVariantsDoNotShareAGroupColor()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestCaseVariantsDoNotShareAGroupColor"
    On Error GoTo Fail

    Dim colors As BetterArray

    Subject.UpdateFromRegistry RegistrySheet, "French"

    'Every spelling appears once, so no group exists even though Excel's own
    'duplicate rule reads the first two as a pair.
    FillLanguageColumn "French", Array("Bonjour", "bonjour", "BONJOUR", "Salut", "Adieu", "Ciao")
    RunUpdateThatAddsARow

    Set colors = GroupRuleColors("French")

    Assert.AreEqual CLng(0), colors.Length, "Two spellings of one word are two labels and neither is a group"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCaseVariantsDoNotShareAGroupColor", Err.Number, Err.Description
End Sub

'@TestMethod("SetupTranslationsTable")
Public Sub TestDuplicateGroupsMatchTheSummary()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestDuplicateGroupsMatchTheSummary"
    On Error GoTo Fail

    Dim colors As BetterArray
    Dim summary As String

    Subject.UpdateFromRegistry RegistrySheet, "French"

    FillLanguageColumn "French", Array("Bonjour", "Salut", "Bonjour", "Salut", "Ciao", "Ciao")
    RunUpdateThatAddsARow

    Set colors = GroupRuleColors("French")
    Assert.IsTrue Subject.DuplicateLabels(summary, "French"), "The three repeated values must be reported"

    Assert.AreEqual CLng(3), colors.Length, "One rule per group the summary names"
    Assert.AreEqual CLng(3), CountOccurrences(summary, " duplicates"), _
                    "The summary names the same three groups the sheet colours"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDuplicateGroupsMatchTheSummary", Err.Number, Err.Description
End Sub


'@TestMethod("SetupTranslationsTable")
'@details The key language is the source of truth every other column is
'         translated from. A repeat there is the sheet's own doing, not
'         something a translator can act on, so it takes no rule at all: not the
'         group fills and not the red catch-all. The translations still take
'         both, which is what the second half of this checks.
Public Sub TestKeyLanguageColumnTakesNoDuplicateRules()
    CustomTestSetTitles Assert, "SetupTranslationsTable", "TestKeyLanguageColumnTakesNoDuplicateRules"
    On Error GoTo Fail

    Dim keyRules As Long
    Dim translationRules As Long

    Subject.UpdateFromRegistry RegistrySheet, "French"
    FillLanguageColumn "French", Array("Salut", "Salut")

    RunUpdateThatAddsARow

    keyRules = ColumnRuleCount("English")
    translationRules = ColumnRuleCount("French")

    Assert.LogSuccesses "TestKeyLanguageColumnTakesNoDuplicateRules: key=" & CStr(keyRules) & _
                        " translation=" & CStr(translationRules)

    Assert.AreEqual CLng(0), keyRules, "The key language column carries no duplicate rule"
    Assert.IsTrue translationRules > 0, "A translation column still carries its duplicate rules"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestKeyLanguageColumnTakesNoDuplicateRules", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================
'@sub-title Write one value per row into a language column, top down.
Private Sub FillLanguageColumn(ByVal columnName As String, ByVal values As Variant)
    Dim dataRange As Range
    Dim idx As Long
    Dim rowIndex As Long

    Set dataRange = TranslationsTable.ListColumns(columnName).DataBodyRange
    If dataRange Is Nothing Then Exit Sub

    rowIndex = 0
    For idx = LBound(values) To UBound(values)
        rowIndex = rowIndex + 1
        If rowIndex > dataRange.Rows.Count Then Exit For
        dataRange.Cells(rowIndex, 1).Value = values(idx)
    Next idx
End Sub

'@sub-title Run an update that brings one new label in, so ApplyFormatting fires.
'@details ApplyFormatting runs only when a row was added or removed. Everything
'         these tests set up is a value edit, which leaves the table clean.
Private Sub RunUpdateThatAddsARow()
    SourceSheet.Range("A3").Value = "Fresh label " & CStr(TranslationsTable.ListRows.Count)
    DropName "RNG_Greetings"
    FixtureWorkbook.Names.Add Name:="RNG_Greetings", RefersTo:=SourceSheet.Range("A1:A3")

    SetRegistryStatus "yes", "yes", "yes"
    Subject.UpdateFromRegistry RegistrySheet
End Sub

'@sub-title Collect the fill of every group rule on a language column.
Private Function GroupRuleColors(ByVal columnName As String) As BetterArray
    Dim dataRange As Range
    Dim idx As Long
    Dim colors As BetterArray

    Set colors = New BetterArray
    colors.LowerBound = 0
    Set GroupRuleColors = colors

    Set dataRange = TranslationsTable.ListColumns(columnName).DataBodyRange
    If dataRange Is Nothing Then Exit Function

    For idx = 1 To dataRange.FormatConditions.Count
        If dataRange.FormatConditions(idx).Type = xlExpression Then
            colors.Push CLng(dataRange.FormatConditions(idx).Interior.Color)
        End If
    Next idx
End Function

'@sub-title Count the conditional formatting rules on a language column.
Private Function ColumnRuleCount(ByVal columnName As String) As Long
    Dim dataRange As Range

    Set dataRange = TranslationsTable.ListColumns(columnName).DataBodyRange
    If dataRange Is Nothing Then Exit Function

    ColumnRuleCount = dataRange.FormatConditions.Count
End Function

'@sub-title Read the fill of the lowest priority rule on a language column.
Private Function LastRuleColor(ByVal columnName As String) As Long
    Dim dataRange As Range
    Dim ruleCount As Long

    Set dataRange = TranslationsTable.ListColumns(columnName).DataBodyRange
    If dataRange Is Nothing Then Exit Function

    ruleCount = dataRange.FormatConditions.Count
    If ruleCount = 0 Then Exit Function

    LastRuleColor = CLng(dataRange.FormatConditions(ruleCount).Interior.Color)
End Function

'@sub-title Test whether a colour is one the palette hands out.
Private Function IsPaletteColor(ByVal candidate As Long) As Boolean
    Dim idx As Long

    For idx = 0 To Subject.DuplicateColorCount - 1
        If Subject.DuplicateGroupColor(idx) = candidate Then
            IsPaletteColor = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Count how many times a marker appears inside a piece of text.
Private Function CountOccurrences(ByVal haystack As String, ByVal needle As String) As Long
    Dim position As Long
    Dim found As Long

    If LenB(needle) = 0 Then Exit Function

    position = InStr(1, haystack, needle, vbBinaryCompare)
    Do While position > 0
        found = found + 1
        position = InStr(position + Len(needle), haystack, needle, vbBinaryCompare)
    Loop

    CountOccurrences = found
End Function


Private Function LastTableRow() As Long
    LastTableRow = TranslationsTable.Range.Row + TranslationsTable.Range.Rows.Count - 1
End Function

Private Function TagHeaderCell() As Range
    Set TagHeaderCell = TranslationsTable.HeaderRowRange.Cells(1, 1).Offset(0, -1)
End Function

Private Function TagDataRange() As Range
    Set TagDataRange = TagHeaderCell().Offset(1, 0).Resize(TranslationsTable.ListRows.Count, 1)
End Function

Private Sub AssertTagBordersAreWhite(ByVal tagData As Range, ByVal context As String)
    AssertOneBorderIsWhite tagData, xlEdgeTop, "top", context
    AssertOneBorderIsWhite tagData, xlEdgeBottom, "bottom", context
    AssertOneBorderIsWhite tagData, xlEdgeLeft, "left", context
    AssertOneBorderIsWhite tagData, xlEdgeRight, "right", context
    AssertOneBorderIsWhite tagData, xlInsideHorizontal, "inside horizontal", context
End Sub

Private Sub AssertOneBorderIsWhite(ByVal tagData As Range, _
                                   ByVal edge As Long, _
                                   ByVal edgeName As String, _
                                   ByVal context As String)
    Assert.AreEqual CLng(vbWhite), CLng(tagData.Borders(edge).Color), _
                    "The " & edgeName & " border of the tag column is white " & context
End Sub

Private Sub AssertSheetSetup()
    EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
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
    matrix = RowsToMatrix(Array( _
        Array("TableName", "rngname", "status", "mode"), _
        Array("table", "RNG_Greetings", "yes", "translate as text"), _
        Array("table", "RNG_Farewell", "no", "translate as text"), _
        Array("table", "RNG_Formula", "yes", "translate as formula")))

    targetSheet.Cells.Clear
    WriteMatrix targetSheet.Cells(1, 1), matrix

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
        If StrComp(CStr(row.Range.Cells(1, 1).Value), label, vbTextCompare) = 0 Then
            TagForLabel = CStr(row.Range.Cells(1, 1).Offset(0, -1).Value)
            Exit Function
        End If
    Next row
End Function

Private Sub ResetTranslationsTableRows()
    Dim attempts As Long

    'The attempt count bounds the loop. A Delete that fails under On Error
    'Resume Next would otherwise spin here for ever.
    On Error Resume Next
        Do While TranslationsTable.ListRows.Count > 0
            TranslationsTable.ListRows(TranslationsTable.ListRows.Count).Delete
            attempts = attempts + 1
            If attempts > 1000 Then Exit Do
        Loop
    On Error GoTo 0
End Sub

Private Sub AppendTranslationLabel(ByVal label As String)
    Dim newRow As ListRow
    Set newRow = TranslationsTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = label
End Sub

'@sub-title Append a row carrying a label and, when supplied, its helper column tag.
Private Sub AppendTaggedLabel(ByVal label As String, ByVal tag As String)
    Dim newRow As ListRow

    Set newRow = TranslationsTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = label

    If LenB(tag) > 0 Then
        newRow.Range.Cells(1, 1).Offset(0, -1).Value = tag
    End If
End Sub

'@sub-title Write count labels into the source sheet and point the registry at them.
Private Sub BuildLargeSource(ByVal count As Long)
    Dim block() As Variant
    Dim idx As Long
    Dim sourceRange As Range

    SourceSheet.Cells.Clear

    ReDim block(1 To count, 1 To 1)
    For idx = 1 To count
        block(idx, 1) = LargeLabel(idx)
    Next idx

    Set sourceRange = SourceSheet.Range(SourceSheet.Cells(1, 1), SourceSheet.Cells(count, 1))
    sourceRange.Value = block

    DropName LARGE_RANGE_NAME
    FixtureWorkbook.Names.Add Name:=LARGE_RANGE_NAME, RefersTo:=sourceRange

    UseSingleRegistryRow LARGE_RANGE_NAME
End Sub

'@sub-title Replace every registry table with one row watching the supplied range.
Private Sub UseSingleRegistryRow(ByVal rangeName As String)
    Dim matrix As Variant

    RemoveRegistryTables
    RegistrySheet.Cells.Clear

    matrix = RowsToMatrix(Array( _
                              Array("colname", "rngname", "updated", "headername"), _
                              Array("table", rangeName, "yes", "translate as text")))
    WriteMatrix RegistrySheet.Cells(1, 1), matrix

    Set RegistryTable = RegistrySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=RegistrySheet.Range("A1:D2"), XlListObjectHasHeaders:=xlYes)
    RegistryTable.Name = REGISTRY_TABLE_NAME
End Sub

Private Sub RemoveRegistryTables()
    Dim idx As Long

    On Error Resume Next
        For idx = RegistrySheet.ListObjects.Count To 1 Step -1
            RegistrySheet.ListObjects(idx).Delete
        Next idx
    On Error GoTo 0

    Set RegistryTable = Nothing
End Sub

'@sub-title Fill the translations table with rowCount labels and their French values.
Private Sub FillTranslationsRows(ByVal rowCount As Long)
    Dim tableRange As Range
    Dim labels() As Variant
    Dim translations() As Variant
    Dim idx As Long

    ResetTranslationsTableRows

    Set tableRange = TranslationsTable.Range
    TranslationsTable.Resize tableRange.Resize(rowCount + 1, tableRange.Columns.Count)

    ReDim labels(1 To rowCount, 1 To 1)
    ReDim translations(1 To rowCount, 1 To 1)
    For idx = 1 To rowCount
        labels(idx, 1) = "Row " & PadNumber(idx)
        translations(idx, 1) = "Fr " & PadNumber(idx)
    Next idx

    TranslationsTable.ListColumns("English").DataBodyRange.Value = labels
    TranslationsTable.ListColumns("French").DataBodyRange.Value = translations
End Sub

'@sub-title Read a language column in one crossing for before and after comparisons.
Private Function SnapshotColumnValues(ByVal columnName As String) As Variant
    Dim target As Range

    On Error Resume Next
        Set target = TranslationsTable.ListColumns(columnName).DataBodyRange
    On Error GoTo 0

    If target Is Nothing Then Exit Function

    SnapshotColumnValues = target.Value2
End Function

Private Function ScreenStateSnapshot() As Boolean
    ScreenStateSnapshot = Application.ScreenUpdating
End Function

Private Sub AssertScreenRestored(ByVal snapshot As Boolean, ByVal routineName As String)
    Assert.AreEqual snapshot, Application.ScreenUpdating, _
                    routineName & " must leave screen updating as it found it"
End Sub

'@sub-title Run an update that is expected to fail and hand back the error it raised.
Private Function RunFailingUpdate() As Long
    On Error Resume Next
        Subject.UpdateFromRegistry RegistrySheet
        RunFailingUpdate = Err.Number
    On Error GoTo 0

    Err.Clear
End Function

'@sub-title Run a language add that is expected to fail and hand back the error it raised.
Private Function RunFailingEnsureLanguages() As Long
    On Error Resume Next
        Subject.EnsureLanguages
        RunFailingEnsureLanguages = Err.Number
    On Error GoTo 0

    Err.Clear
End Function

Private Function LabelExists(ByVal label As String) As Boolean
    LabelExists = (CountLabelIgnoringCase(label) > 0)
End Function

Private Function CountLabelIgnoringCase(ByVal label As String) As Long
    Dim row As ListRow

    For Each row In TranslationsTable.ListRows
        If StrComp(CStr(row.Range.Cells(1, 1).Value), label, vbTextCompare) = 0 Then
            CountLabelIgnoringCase = CountLabelIgnoringCase + 1
        End If
    Next row
End Function

'@sub-title Count the rows spelled exactly this way, case and all.
Private Function CountLabelMatchingCase(ByVal label As String) As Long
    Dim row As ListRow

    For Each row In TranslationsTable.ListRows
        If StrComp(CStr(row.Range.Cells(1, 1).Value), label, vbBinaryCompare) = 0 Then
            CountLabelMatchingCase = CountLabelMatchingCase + 1
        End If
    Next row
End Function

Private Function TranslationForLabel(ByVal label As String, ByVal languageName As String) As String
    Dim labelColumn As Range
    Dim rowIndex As Long

    Set labelColumn = TranslationsTable.ListColumns(1).DataBodyRange
    If labelColumn Is Nothing Then Exit Function

    For rowIndex = 1 To labelColumn.Rows.Count
        If StrComp(CStr(labelColumn.Cells(rowIndex, 1).Value), label, vbTextCompare) = 0 Then
            TranslationForLabel = CStr(TranslationsTable.ListColumns(languageName).DataBodyRange.Cells(rowIndex, 1).Value)
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub DropName(ByVal nameText As String)
    On Error Resume Next
        FixtureWorkbook.Names(nameText).Delete
    On Error GoTo 0
End Sub

Private Function LargeLabel(ByVal idx As Long) As String
    LargeLabel = "Large label " & PadNumber(idx)
End Function

Private Function PadNumber(ByVal idx As Long) As String
    PadNumber = Right$("000" & CStr(idx), 3)
End Function

Private Sub SetCounterValue(ByVal counterValue As Long)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(RegistrySheet)
    If store Is Nothing Then Exit Sub

    store.EnsureName COUNTER_NAME, counterValue, HiddenNameTypeLong
    store.SetValue COUNTER_NAME, counterValue
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
    Dim store As HiddenNames

    Set store = HiddenCounterStore()
    If store Is Nothing Then Exit Function

    CounterValue = store.ValueAsLong(COUNTER_NAME, 0)
End Function

Private Function HiddenCounterExists() As Boolean
    Dim store As HiddenNames

    Set store = HiddenCounterStore()
    If store Is Nothing Then Exit Function

    HiddenCounterExists = store.HasName(COUNTER_NAME)
End Function

Private Function HiddenCounterStore() As HiddenNames
    On Error Resume Next
        Set HiddenCounterStore = HiddenNames.Create(RegistrySheet)
    On Error GoTo 0
End Function
