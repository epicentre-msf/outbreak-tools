Attribute VB_Name = "TestSetupImport"
Option Explicit


'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the improved setup import service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private Service As SetupImport
Private ProgressStub As ProgressDisplayStub
Private PasswordsHandler As Passwords

Private Const PASSWORD_SHEET As String = "TST_SetupImport_Passwords"
Private Const CLEAN_TARGET_SHEET As String = "TST_SetupImport_Clean"
Private Const DICTIONARY_SHEET_NAME As String = "Dictionary"
Private Const EXPORTS_SHEET_NAME As String = "Exports"
Private Const ANALYSIS_SHEET_NAME As String = "Analysis"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const CHOICES_TABLE_NAME As String = "TST_Choices"
Private Const EXPORTS_TABLE_NAME As String = "TST_Exports"
Private Const FOREIGN_PASSWORD As String = "tst_foreign_key"
Private Const MISSING_SHEET_NAME As String = "TST_SetupImport_NoSuchSheet"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const REGISTRY_SOURCE_SHEET As String = "TST_SetupImport_RegistrySource"
Private Const REGISTRY_TABLE_NAME As String = "TST_Registry"
Private Const REGISTRY_RANGE_NAME As String = "RNG_HostMessages"
Private Const REGISTRY_COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const HOST_DICTIONARY_VARIABLE As String = "host_variable"
Private Const SOURCE_DICTIONARY_VARIABLE As String = "import_case_id"
Private Const HOST_EXPORT_STATUS As String = "inactive"
Private Const SOURCE_EXPORT_STATUS As String = "active"
Private Const HOST_EXPORT_LABEL As String = "Host Export"
Private Const SOURCE_EXPORT_LABEL As String = "Imported Export"
Private Const HOST_EXPORT_FILE_NAME As String = "host_export.xlsx"
Private Const SOURCE_EXPORT_FILE_NAME As String = "import_export.xlsx"
Private Const HOST_TRANSLATION_VALUE As String = "Host translation"
Private Const SOURCE_TRANSLATION_VALUE As String = "Imported translation"
Private Const HOST_TRANSLATION_TAG As String = "host_tag"
Private Const SOURCE_TRANSLATION_TAG As String = "import_tag"
Private Const SOURCE_ANALYSIS_HEADER As String = "Analysis imported from workbook"
Private Const DICTIONARY_HOST_START_ROW As Long = 5
Private Const DICTIONARY_HOST_START_COLUMN As Long = 1
Private Const EXPORT_HOST_START_ROW As Long = 4
Private Const EXPORT_HOST_START_COLUMN As Long = 1
Private Const CHOICES_HOST_START_ROW As Long = 4
Private Const CHOICES_HOST_START_COLUMN As Long = 1
Private Const TRANSLATION_HOST_START_ROW As Long = 5
Private Const TRANSLATION_HOST_START_COLUMN As Long = 2
Private Const SOURCE_START_ROW As Long = 1
Private Const SOURCE_START_COLUMN As Long = 1
Private Const TRANSLATION_SOURCE_START_ROW As Long = 1
Private Const TRANSLATION_SOURCE_START_COLUMN As Long = 2
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private KeepExportArtifacts As Boolean
Private SharedExportBook As Workbook
Private SharedExportPath As String

'@section Module lifecycle
'===============================================================================
'The four hooks below each carry a handler. The runner calls them through
'Application.Run inside an Apple Events call, and a failure with no handler
'reaches the VBE as a dialog that takes the whole run down with an opaque -50.

'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail

    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupImport"
    KeepExportArtifacts = False
    Exit Sub

Fail:
    Err.Clear
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        ReleaseSharedExport
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail

    Set ProgressStub = New ProgressDisplayStub
    ProgressStub.Caption = vbNullString
    ProgressStub.Value = vbNullString
    Set Service = New SetupImport
    Service.Path = ThisWorkbook.FullName
    Set Service.ProgressObject = ProgressStub
    EnsurePasswordsFixture
    Exit Sub

Fail:
    'The test that follows reports the real failure through its own handler.
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Service = Nothing
    Set ProgressStub = Nothing
    Set PasswordsHandler = Nothing
    DeleteWorksheet CLEAN_TARGET_SHEET
    DeleteWorksheet PASSWORD_SHEET
    DeleteWorksheet REGISTRY_SHEET_NAME
    DeleteWorksheet REGISTRY_SOURCE_SHEET
    DeleteWorksheet CHOICES_SHEET_NAME
    DeleteWorksheet DICTIONARY_SHEET_NAME
    DeleteWorksheet EXPORTS_SHEET_NAME
    DeleteWorksheet ANALYSIS_SHEET_NAME
    DeleteWorksheet TRANSLATIONS_SHEET_NAME
    ThisWorkbook.Names(REGISTRY_RANGE_NAME).Delete
    On Error GoTo 0
End Sub


'@section Tests - Check and Clean
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestCheckRaisesWhenNoSelection()
    CustomTestSetTitles Assert, "SetupImport", "TestCheckRaisesWhenNoSelection"
    On Error GoTo ExpectInvalid

    Service.Check False, False, False, False, False
    Assert.LogFailure "Check should raise when no import option is selected."
    Exit Sub

ExpectInvalid:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Unexpected error code."
    Assert.AreEqual "Please select at least one import option (Dictionary, Choices, Exports, Analysis or Translations).", _
                    ProgressStub.Value, "Expected message to be surfaced through the progress display."
    Assert.AreEqual ProgressStub.Value, ProgressStub.Caption, "Caption should mirror value for progress updates."
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestCheckRaisesWhenFileMissing()
    CustomTestSetTitles Assert, "SetupImport", "TestCheckRaisesWhenFileMissing"
    Dim missingPath As String

    On Error GoTo Fail
        missingPath = BuildMissingSetupPath()
        Service.Path = missingPath

    On Error GoTo ExpectMissing
        Service.Check True, False, False, False, False
        Assert.LogFailure "Check should raise when the source workbook cannot be located."
        Exit Sub

ExpectMissing:
    Assert.AreEqual CLng(ProjectError.ElementNotFound), Err.Number, "Unexpected error code when file is missing."
    Assert.IsTrue InStr(1, ProgressStub.Value, missingPath, vbTextCompare) > 0, _
                   "Progress display should include the missing path."
    Assert.IsTrue InStr(1, ProgressStub.Caption, missingPath, vbTextCompare) > 0, _
                   "Caption should also include the missing path."
    Err.Clear
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCheckRaisesWhenFileMissing", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestCleanRemovesWorksheetComments()
    CustomTestSetTitles Assert, "SetupImport", "TestCleanRemovesWorksheetComments"
    Dim targetSheet As Worksheet
    Dim sheetsList As BetterArray

    On Error GoTo Fail

    Set targetSheet = EnsureWorksheet(CLEAN_TARGET_SHEET)
    PrepareComment targetSheet

    Set sheetsList = SheetsListOf(CLEAN_TARGET_SHEET)
    Service.Clean PasswordsHandler, sheetsList

    Assert.IsTrue targetSheet.Cells(1, 1).Comment Is Nothing, "Clean should remove classic comments."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCleanRemovesWorksheetComments", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Tests - Import
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestImportClosesWorkbookAfterRun()
    CustomTestSetTitles Assert, "SetupImport", "TestImportClosesWorkbookAfterRun"
    Dim tempBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim workbookName As String

    On Error GoTo CleanupFailure

    Set tempBook = NewWorkbook
    tempBook.Worksheets(1).Name = "TempData"

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupImportTests")
    workbookPath = BuildWorkbookPath(exportFolder, "setup_import_source", ".xlsx")
    tempBook.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    tempBook.Close SaveChanges:=False

    workbookName = FileNameFromPath(workbookPath)
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf("MissingSheet")

    Service.Import PasswordsHandler, sheetsList
    Assert.IsFalse IsWorkbookOpen(workbookName), "Import should close the source workbook on completion."

    'Calling Import again should reopen and close the workbook without errors.
    Service.Import PasswordsHandler, sheetsList
    Assert.IsFalse IsWorkbookOpen(workbookName), "Import should leave no lingering workbook reference."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportClosesWorkbookAfterRun", workbookPath
End Sub

'@TestMethod("SetupImport")
Public Sub TestImportFromWorkbookUsingDomainClasses()
    CustomTestSetTitles Assert, "SetupImport", "TestImportFromWorkbookUsingDomainClasses"
    Dim sourceBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim workbookName As String
    Dim sheetsList As BetterArray

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Set sourceBook = BuildImportWorkbookFixture()
    exportFolder = BuildTempFolder(ThisWorkbook, "SetupImportDomainTests")
    workbookPath = BuildWorkbookPath(exportFolder, "setup_import_domain", ".xlsx")

    sourceBook.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    workbookName = FileNameFromPath(workbookPath)
    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing

    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(DICTIONARY_SHEET_NAME, EXPORTS_SHEET_NAME, ANALYSIS_SHEET_NAME, TRANSLATIONS_SHEET_NAME)

    Service.ImportFromWorkbook PasswordsHandler, sheetsList

    AssertImportedDictionary
    AssertImportedExports
    AssertImportedAnalysis
    AssertImportedTranslations

    Assert.IsFalse IsWorkbookOpen(workbookName), "ImportFromWorkbook should close the source workbook."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    DeleteFileIfExists workbookPath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestImportFromWorkbookUsingDomainClasses", errNumber, errDescription
        Err.Clear
    End If
    Exit Sub
End Sub


'@TestMethod("SetupImport")
Public Sub TestImportMarksTranslationsForReview()
    CustomTestSetTitles Assert, "SetupImport", "TestImportMarksTranslationsForReview"
    Dim sourceBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim translationSheet As Worksheet
    Dim firstTag As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Set sourceBook = BuildImportWorkbookFixture()
    exportFolder = BuildTempFolder(ThisWorkbook, "SetupImportReviewTests")
    workbookPath = BuildWorkbookPath(exportFolder, "setup_import_review", ".xlsx")

    sourceBook.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing

    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(DICTIONARY_SHEET_NAME, TRANSLATIONS_SHEET_NAME)

    'The sheet-list path, not ImportFromWorkbook: both must leave the review behind.
    Service.Import PasswordsHandler, sheetsList

    Set translationSheet = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
    firstTag = CStr(translationSheet.Cells(TRANSLATION_HOST_START_ROW + 1, TRANSLATION_HOST_START_COLUMN - 1).Value)
    Assert.AreEqual "__imported____0", firstTag, _
                    "Import should mark every imported translation row for review."
    Assert.IsTrue HiddenNames.Create(translationSheet).HasName("__SetupTranslationsUnseenReview__"), _
                  "Import should ask the next update to review the unseen labels."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    DeleteFileIfExists workbookPath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestImportMarksTranslationsForReview", errNumber, errDescription
        Err.Clear
    End If
    Exit Sub
End Sub
'@TestMethod("SetupImport")
Public Sub TestImportFromWorkbookSkipsMissingSheets()
    CustomTestSetTitles Assert, "SetupImport", "TestImportFromWorkbookSkipsMissingSheets"
    Dim sourceBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim sheetsList As BetterArray

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Set sourceBook = BuildImportWorkbookFixture()
    On Error Resume Next
        sourceBook.Worksheets(TRANSLATIONS_SHEET_NAME).Delete
    On Error GoTo 0

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupImportDomainTests")
    workbookPath = BuildWorkbookPath(exportFolder, "setup_import_missing", ".xlsx")

    sourceBook.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing

    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(DICTIONARY_SHEET_NAME, TRANSLATIONS_SHEET_NAME)

    Service.ImportFromWorkbook PasswordsHandler, sheetsList

    AssertImportedDictionary
    AssertTranslationUnchanged

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    DeleteFileIfExists workbookPath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestImportFromWorkbookSkipsMissingSheets", errNumber, errDescription
        Err.Clear
    End If
    Exit Sub
End Sub


'@section Tests - Export cancellation and file creation
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestExportAbortsWhenFolderSelectionCancelled()
    CustomTestSetTitles Assert, "SetupImport", "TestExportAbortsWhenFolderSelectionCancelled"
    Dim initialWorkbookCount As Long
    Dim svc As SetupImport

    On Error GoTo Fail

    PrepareHostSetupSheets

    Service.DisplayPrompts = False
    Service.SetExportFolder vbNullString

    initialWorkbookCount = Application.Workbooks.Count
    Set svc = Service
    svc.Export

    Assert.AreEqual initialWorkbookCount, Application.Workbooks.Count, _
                     "Export should not create workbooks when no folder is selected."
    Assert.AreEqual vbNullString, svc.LastExportFile, _
                     "Export should not record a file path when cancelled."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportAbortsWhenFolderSelectionCancelled", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportCreatesWorkbookInProvidedFolder()
    CustomTestSetTitles Assert, "SetupImport", "TestExportCreatesWorkbookInProvidedFolder"
    Dim exportFolder As String
    Dim expectedFilePath As String
    Dim expectedPrefix As String
    Dim svc As SetupImport
    Dim initialWorkbookCount As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Fail

    PrepareHostSetupSheets

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupExportTests")

    Service.DisplayPrompts = False
    Service.SetExportFolder exportFolder

    initialWorkbookCount = Application.Workbooks.Count
    Set svc = Service
    svc.Export

    expectedFilePath = svc.LastExportFile
    expectedPrefix = exportFolder & Application.PathSeparator & HostBaseName() & "_export_"

    Assert.IsTrue LenB(expectedFilePath) > 0, "Export should expose the saved file path."
    Assert.IsTrue LenB(Dir$(expectedFilePath)) > 0, "Export should write the workbook to the configured folder."
    Assert.AreEqual initialWorkbookCount, Application.Workbooks.Count, "Export should close the temporary export workbook."
    Assert.AreEqual expectedPrefix, Left$(expectedFilePath, Len(expectedPrefix)), _
                    "Export should write into the configured folder under the host name."

    If Not KeepExportArtifacts Then
        DeleteFileIfExists expectedFilePath
    End If
    Exit Sub

Fail:
    'The error is read before DeleteFileIfExists, because that call clears Err.
    errNumber = Err.Number
    errDescription = Err.Description
    If Not KeepExportArtifacts Then
        DeleteFileIfExists expectedFilePath
    End If
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportCreatesWorkbookInProvidedFolder", errNumber, errDescription
        Err.Clear
    End If
End Sub


'@section Tests - Export component verification
'===============================================================================

'@TestMethod("SetupImport")
Public Sub TestExportContainsDictionarySheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsDictionarySheet"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."
    Assert.IsTrue ExportWorksheetExists(exportBook, DICTIONARY_SHEET_NAME), _
                  "Export workbook should contain a Dictionary worksheet."

    Set exportedSheet = exportBook.Worksheets(DICTIONARY_SHEET_NAME)

    'Exported dictionary starts at row 1 with headers
    Assert.AreEqual "variable name", LCase$(CStr(exportedSheet.Cells(1, 1).Value)), _
                    "Dictionary export should place Variable Name as the first header."

    'Verify host variable data is present in the first data row
    Assert.AreEqual HOST_DICTIONARY_VARIABLE, CStr(exportedSheet.Cells(2, 1).Value), _
                    "Dictionary export should include the host variable name in the first data row."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsDictionarySheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportContainsChoicesSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsChoicesSheet"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim labelColumn As Long

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."
    Assert.IsTrue ExportWorksheetExists(exportBook, CHOICES_SHEET_NAME), _
                  "Export workbook should contain a Choices worksheet."

    Set exportedSheet = exportBook.Worksheets(CHOICES_SHEET_NAME)

    'Verify choices headers are present at row 1
    Assert.AreEqual "list name", LCase$(CStr(exportedSheet.Cells(1, 1).Value)), _
                    "Choices export should place list name as the first header."

    'Verify choices data is present (list_primary from the fixture)
    Assert.AreEqual "list_primary", CStr(exportedSheet.Cells(2, 1).Value), _
                    "Choices export should include the first choice list name in the data."

    'The setup Choices sheet carries six columns and "label" is the fifth of
    'them, behind "non translated label" and "translated label". The column is
    'found by its header, so a layout change reports as a missing column rather
    'than as wrong data.
    labelColumn = ExportedHeaderColumn(exportedSheet, "label")
    Assert.IsTrue labelColumn > 0, "Choices export should carry a label column."
    If labelColumn > 0 Then
        Assert.AreEqual "Choice A", CStr(exportedSheet.Cells(2, labelColumn).Value), _
                        "Choices export should include the label column data."
    End If
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsChoicesSheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportContainsExportSpecsSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsExportSpecsSheet"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim lastColumn As Long
    Dim colIdx As Long
    Dim statusColumn As Long
    Dim fileNameColumn As Long
    Dim headerValue As String

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."
    Assert.IsTrue ExportWorksheetExists(exportBook, EXPORTS_SHEET_NAME), _
                  "Export workbook should contain an Exports worksheet."

    Set exportedSheet = exportBook.Worksheets(EXPORTS_SHEET_NAME)

    'Find the status and file name columns by scanning headers
    lastColumn = exportedSheet.Cells(1, exportedSheet.Columns.Count).End(xlToLeft).Column
    statusColumn = 0
    fileNameColumn = 0

    For colIdx = 1 To lastColumn
        headerValue = LCase$(CStr(exportedSheet.Cells(1, colIdx).Value))
        If headerValue = "status" Then statusColumn = colIdx
        If headerValue = "file name" Then fileNameColumn = colIdx
    Next colIdx

    Assert.IsTrue statusColumn > 0, _
                  "Exports export should contain a status column header."
    Assert.IsTrue fileNameColumn > 0, _
                  "Exports export should contain a file name column header."

    'Verify data values match the host fixture
    If statusColumn > 0 Then
        Assert.AreEqual HOST_EXPORT_STATUS, CStr(exportedSheet.Cells(2, statusColumn).Value), _
                        "Exports export should carry the host export status value."
    End If
    If fileNameColumn > 0 Then
        Assert.AreEqual HOST_EXPORT_FILE_NAME, CStr(exportedSheet.Cells(2, fileNameColumn).Value), _
                        "Exports export should carry the host export file name."
    End If
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsExportSpecsSheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportContainsAnalysisSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsAnalysisSheet"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim summaryTable As ListObject

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."
    Assert.IsTrue ExportWorksheetExists(exportBook, ANALYSIS_SHEET_NAME), _
                  "Export workbook should contain an Analysis worksheet."

    Set exportedSheet = exportBook.Worksheets(ANALYSIS_SHEET_NAME)

    'Verify at least one analysis table was exported
    Assert.IsTrue exportedSheet.ListObjects.Count > 0, _
                  "Analysis export should contain at least one ListObject."

    'Verify the global summary table exists with correct data
    On Error Resume Next
        Set summaryTable = exportedSheet.ListObjects("Tab_global_summary")
    On Error GoTo 0

    Assert.IsTrue Not (summaryTable Is Nothing), _
                  "Analysis export should contain the Tab_global_summary table."
    If Not summaryTable Is Nothing Then
        Assert.AreEqual "Host global section", _
                        CStr(summaryTable.DataBodyRange.Cells(1, 1).Value), _
                        "Analysis export should preserve the global summary data."
    End If
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsAnalysisSheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportContainsMultipleAnalysisTables()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsMultipleAnalysisTables"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim univariateTable As ListObject
    Dim timeseriesTable As ListObject

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(ANALYSIS_SHEET_NAME)

    'Verify multiple analysis tables are exported (not just the first one)
    On Error Resume Next
        Set univariateTable = exportedSheet.ListObjects("Tab_Univariate_Analysis")
        Set timeseriesTable = exportedSheet.ListObjects("Tab_TimeSeries_Analysis")
    On Error GoTo 0

    Assert.IsTrue Not (univariateTable Is Nothing), _
                  "Analysis export should include the Univariate Analysis table."
    Assert.IsTrue Not (timeseriesTable Is Nothing), _
                  "Analysis export should include the TimeSeries Analysis table."

    If Not timeseriesTable Is Nothing Then
        Assert.IsTrue timeseriesTable.ListRows.Count >= 2, _
                      "TimeSeries table should contain at least two data rows."
    End If
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsMultipleAnalysisTables", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Tests - Export structural checks
'===============================================================================

'@TestMethod("SetupImport")
Public Sub TestExportRemovesDefaultWorksheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportRemovesDefaultWorksheet"
    Dim exportBook As Workbook

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    'The default Sheet1 from Workbooks.Add should have been removed
    Assert.IsFalse ExportWorksheetExists(exportBook, "Sheet1"), _
                   "Export should remove the default Sheet1 worksheet."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportRemovesDefaultWorksheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportMainSheetsAreVisible()
    CustomTestSetTitles Assert, "SetupImport", "TestExportMainSheetsAreVisible"
    Dim exportBook As Workbook

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    'All main setup sheets should be visible (Hide:=xlSheetVisible is passed in Export)
    If ExportWorksheetExists(exportBook, DICTIONARY_SHEET_NAME) Then
        Assert.AreEqual CLng(xlSheetVisible), CLng(exportBook.Worksheets(DICTIONARY_SHEET_NAME).Visible), _
                        "Dictionary worksheet should be visible in the export."
    End If

    If ExportWorksheetExists(exportBook, CHOICES_SHEET_NAME) Then
        Assert.AreEqual CLng(xlSheetVisible), CLng(exportBook.Worksheets(CHOICES_SHEET_NAME).Visible), _
                        "Choices worksheet should be visible in the export."
    End If

    If ExportWorksheetExists(exportBook, EXPORTS_SHEET_NAME) Then
        Assert.AreEqual CLng(xlSheetVisible), CLng(exportBook.Worksheets(EXPORTS_SHEET_NAME).Visible), _
                        "Exports worksheet should be visible in the export."
    End If

    If ExportWorksheetExists(exportBook, ANALYSIS_SHEET_NAME) Then
        Assert.AreEqual CLng(xlSheetVisible), CLng(exportBook.Worksheets(ANALYSIS_SHEET_NAME).Visible), _
                        "Analysis worksheet should be visible in the export."
    End If

    If ExportWorksheetExists(exportBook, TRANSLATIONS_SHEET_NAME) Then
        Assert.AreEqual CLng(xlSheetVisible), CLng(exportBook.Worksheets(TRANSLATIONS_SHEET_NAME).Visible), _
                        "Translations worksheet should be visible in the export."
    End If
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportMainSheetsAreVisible", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportClosesWorkbookAfterCompletion()
    CustomTestSetTitles Assert, "SetupImport", "TestExportClosesWorkbookAfterCompletion"
    Dim exportFilePath As String
    Dim initialCount As Long
    Dim svc As SetupImport

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Service.DisplayPrompts = False
    Service.SetExportFolder BuildTempFolder(ThisWorkbook, "SetupExportTests")

    initialCount = Application.Workbooks.Count
    Set svc = Service
    svc.Export

    exportFilePath = svc.LastExportFile

    Assert.AreEqual initialCount, Application.Workbooks.Count, _
                     "Export should not leave any open workbooks behind."

    If Not KeepExportArtifacts Then
        DeleteFileIfExists exportFilePath
    End If
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    If Not KeepExportArtifacts Then
        DeleteFileIfExists exportFilePath
    End If
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportClosesWorkbookAfterCompletion", errNumber, errDescription
        Err.Clear
    End If
End Sub


'@section Tests - Export edge cases
'===============================================================================

'@TestMethod("SetupImport")
Public Sub TestExportSkipsMissingHostDictionarySheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportSkipsMissingHostDictionarySheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Dictionary sheet from the host before exporting
    DeleteWorksheet DICTIONARY_SHEET_NAME

    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should succeed even when Dictionary sheet is missing from host."

    'Dictionary should not appear in the export since it was removed from host
    Assert.IsFalse ExportWorksheetExists(exportBook, DICTIONARY_SHEET_NAME), _
                   "Export should not create a Dictionary sheet when the host does not have one."

    'Other sheets should still be exported
    Assert.IsTrue ExportWorksheetExists(exportBook, CHOICES_SHEET_NAME), _
                  "Choices should still be exported when Dictionary is missing."
    Assert.IsTrue ExportWorksheetExists(exportBook, EXPORTS_SHEET_NAME), _
                  "Exports should still be exported when Dictionary is missing."
    Assert.IsTrue ExportWorksheetExists(exportBook, ANALYSIS_SHEET_NAME), _
                  "Analysis should still be exported when Dictionary is missing."
    Assert.IsTrue ExportWorksheetExists(exportBook, TRANSLATIONS_SHEET_NAME), _
                  "Translations should still be exported when Dictionary is missing."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportSkipsMissingHostDictionarySheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportSkipsMissingHostAnalysisSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportSkipsMissingHostAnalysisSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Analysis sheet from the host before exporting
    DeleteWorksheet ANALYSIS_SHEET_NAME

    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should succeed even when Analysis sheet is missing from host."

    'Analysis should not appear in the export
    Assert.IsFalse ExportWorksheetExists(exportBook, ANALYSIS_SHEET_NAME), _
                   "Export should not create an Analysis sheet when the host does not have one."

    'Other sheets should still be present
    Assert.IsTrue ExportWorksheetExists(exportBook, DICTIONARY_SHEET_NAME), _
                  "Dictionary should still be exported when Analysis is missing."
    Assert.IsTrue ExportWorksheetExists(exportBook, CHOICES_SHEET_NAME), _
                  "Choices should still be exported when Analysis is missing."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportSkipsMissingHostAnalysisSheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportSkipsMissingHostChoicesSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportSkipsMissingHostChoicesSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Choices sheet from the host before exporting
    DeleteWorksheet CHOICES_SHEET_NAME

    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should succeed even when Choices sheet is missing from host."

    Assert.IsFalse ExportWorksheetExists(exportBook, CHOICES_SHEET_NAME), _
                   "Export should not create a Choices sheet when the host does not have one."

    'Other sheets should still be present
    Assert.IsTrue ExportWorksheetExists(exportBook, DICTIONARY_SHEET_NAME), _
                  "Dictionary should still be exported when Choices is missing."
    Assert.IsTrue ExportWorksheetExists(exportBook, EXPORTS_SHEET_NAME), _
                  "Exports should still be exported when Choices is missing."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportSkipsMissingHostChoicesSheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportDictionaryContainsListObject()
    CustomTestSetTitles Assert, "SetupImport", "TestExportDictionaryContainsListObject"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(DICTIONARY_SHEET_NAME)

    'LLdictionary.Export creates a ListObject in the exported sheet
    Assert.IsTrue exportedSheet.ListObjects.Count > 0, _
                  "Dictionary export should contain a ListObject wrapping the data."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportDictionaryContainsListObject", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportDictionaryPreservesMultipleHeaders()
    CustomTestSetTitles Assert, "SetupImport", "TestExportDictionaryPreservesMultipleHeaders"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(DICTIONARY_SHEET_NAME)

    'Verify several key dictionary headers are present in the export
    Assert.AreEqual "main label", LCase$(CStr(exportedSheet.Cells(1, 2).Value)), _
                    "Dictionary export should include Main Label as the second header."
    Assert.AreEqual "status", LCase$(CStr(exportedSheet.Cells(1, 11).Value)), _
                    "Dictionary export should include Status header."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportDictionaryPreservesMultipleHeaders", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportChoicesPreservesAllDataRows()
    CustomTestSetTitles Assert, "SetupImport", "TestExportChoicesPreservesAllDataRows"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim lastDataRow As Long

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(CHOICES_SHEET_NAME)

    'The fixture creates 4 choice rows (2 for list_primary, 2 for list_secondary)
    lastDataRow = exportedSheet.Cells(exportedSheet.Rows.Count, 1).End(xlUp).Row

    'Row 1 = headers, rows 2-5 = data (4 rows)
    Assert.IsTrue lastDataRow >= 5, _
                  "Choices export should contain all fixture data rows (at least 4 data rows)."

    'Verify the second list is also present
    Dim foundSecondary As Boolean
    Dim rowIdx As Long

    foundSecondary = False
    For rowIdx = 2 To lastDataRow
        If CStr(exportedSheet.Cells(rowIdx, 1).Value) = "list_secondary" Then
            foundSecondary = True
            Exit For
        End If
    Next rowIdx

    Assert.IsTrue foundSecondary, _
                  "Choices export should include all choice lists from the host."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportChoicesPreservesAllDataRows", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportExportsHiddenNames()
    CustomTestSetTitles Assert, "SetupImport", "TestExportExportsHiddenNames"
    Dim exportBook As Workbook
    Dim exportedName As Name

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    'HiddenNames.ExportNamesToWorkbook should copy named ranges to the export workbook
    'The host has __ll_exports_total__ set to 1
    On Error Resume Next
        Set exportedName = exportBook.Names("__ll_exports_total__")
    On Error GoTo 0

    Assert.IsTrue Not (exportedName Is Nothing), _
                  "Export should include the __ll_exports_total__ hidden name in the workbook."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportExportsHiddenNames", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportTranslationsStartsAtColumnTwo()
    CustomTestSetTitles Assert, "SetupImport", "TestExportTranslationsStartsAtColumnTwo"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(TRANSLATIONS_SHEET_NAME)

    'Translations are exported starting at column 2 to preserve the tag offset
    Assert.AreEqual vbNullString, Trim$(CStr(exportedSheet.Cells(1, 1).Value)), _
                    "Translations export column 1 should be empty (tag offset preserved)."
    Assert.AreEqual "lang1", LCase$(CStr(exportedSheet.Cells(1, 2).Value)), _
                    "Translations export should start headers at column 2."
    Assert.AreEqual "english", LCase$(CStr(exportedSheet.Cells(1, 3).Value)), _
                    "Translations export should carry the English column one column further right."
    Assert.AreEqual HOST_TRANSLATION_VALUE, CStr(exportedSheet.Cells(2, 3).Value), _
                    "Translations export should keep the host translation after the shift."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportTranslationsStartsAtColumnTwo", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Tests - the progress display
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestProgressObjectRejectsUnsupportedDisplay()
    CustomTestSetTitles Assert, "SetupImport", "TestProgressObjectRejectsUnsupportedDisplay"
    On Error GoTo ExpectInvalid

    'A Collection carries neither Caption nor Value.
    Set Service.ProgressObject = New Collection
    Assert.LogFailure "Assigning an object with no Caption and no Value should raise."
    Exit Sub

ExpectInvalid:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Unexpected error code."
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestImportReportsProgress()
    CustomTestSetTitles Assert, "SetupImport", "TestImportReportsProgress"
    Dim workbookPath As String
    Dim sheetsList As BetterArray

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    PrepareHostChoicesTable

    workbookPath = BuildSetupSourceFile("setup_import_progress")
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(CHOICES_SHEET_NAME)

    Service.Import PasswordsHandler, sheetsList

    Assert.IsTrue LenB(ProgressStub.Caption) > 0, "Import should write to the progress display."
    Assert.IsTrue InStr(1, ProgressStub.Caption, "%", vbTextCompare) > 0, _
                  "Import should end with a progress reading on the display."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportReportsProgress", workbookPath
End Sub


'@section Tests - the Choices column round trip
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestImportRestoresChoicesColumnNames()
    CustomTestSetTitles Assert, "SetupImport", "TestImportRestoresChoicesColumnNames"
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim lo As ListObject

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    PrepareHostChoicesTable

    workbookPath = BuildSetupSourceFile("setup_import_choices")
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(CHOICES_SHEET_NAME)

    Service.Import PasswordsHandler, sheetsList

    Set lo = HostChoicesTable()
    Assert.IsTrue Not (lo Is Nothing), "The host Choices table should still be there."
    Assert.AreEqual CLng(1), HeaderCount(lo, "Label"), "Choices should hold one Label column."
    Assert.AreEqual CLng(1), HeaderCount(lo, "Translated Label"), "Choices should hold one Translated Label column."
    Assert.AreEqual CLng(0), HeaderCount(lo, "Formula Label"), _
                    "Formula Label is a name the import borrows and must not survive it."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportRestoresChoicesColumnNames", workbookPath
End Sub

'@TestMethod("SetupImport")
Public Sub TestImportRestoresChoicesColumnNamesAfterFailure()
    CustomTestSetTitles Assert, "SetupImport", "TestImportRestoresChoicesColumnNamesAfterFailure"
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim lo As ListObject
    Dim importRaised As Boolean

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    PrepareHostChoicesTable

    workbookPath = BuildSetupSourceFile("setup_import_choices_fail")
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(CHOICES_SHEET_NAME, EXPORTS_SHEET_NAME)

    'The Exports sheet is locked with a key the handler does not hold, so the
    'write into its table raises part way through the run.
    LockSheetWithForeignKey EXPORTS_SHEET_NAME

    On Error Resume Next
        Service.Import PasswordsHandler, sheetsList
        importRaised = (Err.Number <> 0)
        Err.Clear
    On Error GoTo CleanupFailure

    UnlockForeignKeySheet EXPORTS_SHEET_NAME

    Assert.IsTrue importRaised, "The locked Exports sheet should make the import raise."

    Set lo = HostChoicesTable()
    Assert.IsTrue Not (lo Is Nothing), "The host Choices table should still be there."
    Assert.AreEqual CLng(1), HeaderCount(lo, "Label"), "A failed import should still leave one Label column."
    Assert.AreEqual CLng(1), HeaderCount(lo, "Translated Label"), _
                    "A failed import should still leave one Translated Label column."
    Assert.AreEqual CLng(0), HeaderCount(lo, "Formula Label"), _
                    "A failed import should still take the borrowed name back."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    'LogUnexpected reads Err first, so it comes before anything that clears it.
    LogUnexpected "TestImportRestoresChoicesColumnNamesAfterFailure", workbookPath
    UnlockForeignKeySheet EXPORTS_SHEET_NAME
End Sub

'@TestMethod("SetupImport")
Public Sub TestImportLeavesChoicesProtectionLocked()
    CustomTestSetTitles Assert, "SetupImport", "TestImportLeavesChoicesProtectionLocked"
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim choicesSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    PrepareHostChoicesTable

    workbookPath = BuildSetupSourceFile("setup_import_choices_lock")
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(CHOICES_SHEET_NAME)

    Service.Import PasswordsHandler, sheetsList

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET_NAME)

    'Assert the protection is live first. Passwords.ProtectSheet returns early
    'in debug mode, and the two readings below would then say nothing.
    Assert.IsTrue choicesSheet.ProtectContents, "Choices should end the import protected."
    Assert.IsTrue choicesSheet.ProtectDrawingObjects, "Choices should end the import closed to shape edits."
    Assert.IsFalse choicesSheet.Protection.AllowDeletingRows, "Choices should end the import closed to row deletion."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportLeavesChoicesProtectionLocked", workbookPath
End Sub


'@section Tests - the export row sync
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestImportGrowsHostExportsToMatchSource()
    CustomTestSetTitles Assert, "SetupImport", "TestImportGrowsHostExportsToMatchSource"
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim hostExports As ListObject

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    workbookPath = BuildSetupSourceFile("setup_import_exports", extraExportRows:=1)
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf(EXPORTS_SHEET_NAME)

    Service.Import PasswordsHandler, sheetsList

    Set hostExports = HostExportsTable()
    Assert.IsTrue Not (hostExports Is Nothing), "The host Exports table should still be there."
    Assert.AreEqual CLng(2), CLng(hostExports.ListRows.Count), _
                    "The host exports table should carry as many rows as the setup file."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportGrowsHostExportsToMatchSource", workbookPath
End Sub

'@TestMethod("SetupImport")
Public Sub TestImportFromWorkbookSyncsExportsWithoutSheetList()
    CustomTestSetTitles Assert, "SetupImport", "TestImportFromWorkbookSyncsExportsWithoutSheetList"
    Dim workbookPath As String
    Dim workbookName As String
    Dim hostExports As ListObject

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Dictionary and Exports only. PrepareImport reads both of them, and the
    'domain import leaves the other three sheets alone when the source has none.
    workbookPath = BuildSetupSourceFile("setup_import_nolist", extraExportRows:=1, exportsOnly:=True)
    workbookName = FileNameFromPath(workbookPath)
    Service.Path = workbookPath

    'No sheet list. PrepareImport used to exit on the raw argument and skip the
    'export sync for the whole run.
    Service.ImportFromWorkbook PasswordsHandler

    Set hostExports = HostExportsTable()
    Assert.IsTrue Not (hostExports Is Nothing), "The host Exports table should still be there."
    Assert.AreEqual CLng(2), CLng(hostExports.ListRows.Count), _
                    "A run with no sheet list should still bring the host exports up to the setup file."
    Assert.IsFalse IsWorkbookOpen(workbookName), "ImportFromWorkbook should close the source workbook."

    DeleteFileIfExists workbookPath
    Exit Sub

CleanupFailure:
    LogUnexpected "TestImportFromWorkbookSyncsExportsWithoutSheetList", workbookPath
End Sub


'@section Tests - Clean on a missing sheet
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestCleanSkipsMissingWorksheet()
    CustomTestSetTitles Assert, "SetupImport", "TestCleanSkipsMissingWorksheet"
    Dim targetSheet As Worksheet
    Dim sheetsList As BetterArray

    On Error GoTo Fail

    Set targetSheet = EnsureWorksheet(CLEAN_TARGET_SHEET)
    PrepareComment targetSheet

    'Import skips a sheet the host does not carry. Clean now does the same.
    Set sheetsList = SheetsListOf(MISSING_SHEET_NAME, CLEAN_TARGET_SHEET)
    Service.Clean PasswordsHandler, sheetsList

    Assert.IsTrue targetSheet.Cells(1, 1).Comment Is Nothing, _
                  "Clean should reach the sheets that are there after skipping one that is not."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCleanSkipsMissingWorksheet", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Tests - the export file
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestExportFileNameDropsHostExtension()
    CustomTestSetTitles Assert, "SetupImport", "TestExportFileNameDropsHostExtension"
    Dim fileName As String

    On Error GoTo Fail

    fileName = FileNameFromPath(SharedExportFilePath())

    Assert.IsTrue LenB(fileName) > 0, "Export should produce a file."
    Assert.AreEqual CLng(0), CLng(InStr(1, fileName, ".xlsb", vbTextCompare)), _
                    "The host extension should not travel inside the export name."
    Assert.AreEqual CLng(1), CLng(InStr(1, fileName, HostBaseName() & "_export_", vbTextCompare)), _
                    "The export name should open with the host name."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameDropsHostExtension", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportFileNamesAreUniquePerRun()
    CustomTestSetTitles Assert, "SetupImport", "TestExportFileNamesAreUniquePerRun"
    Dim exportFolder As String
    Dim firstPath As String
    Dim secondPath As String
    Dim svc As SetupImport

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupExportTests")
    Service.DisplayPrompts = False
    Set svc = Service

    Service.SetExportFolder exportFolder
    svc.Export
    firstPath = svc.LastExportFile

    'The stamp counts to the second, and the folder is consumed by one export.
    WaitForNextSecond
    Service.SetExportFolder exportFolder
    svc.Export
    secondPath = svc.LastExportFile

    Assert.IsTrue LenB(firstPath) > 0, "The first export should record a file path."
    Assert.IsTrue LenB(secondPath) > 0, "The second export should record a file path."
    Assert.AreNotEqual firstPath, secondPath, "Two exports in one session should be two files."
    Assert.IsTrue LenB(Dir$(firstPath)) > 0, "The first export should survive the second."

    DeleteFileIfExists firstPath
    DeleteFileIfExists secondPath
    Exit Sub

CleanupFailure:
    DeleteFileIfExists firstPath
    DeleteFileIfExists secondPath
    CustomTestLogFailure Assert, "TestExportFileNamesAreUniquePerRun", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportSavesAsOpenXml()
    CustomTestSetTitles Assert, "SetupImport", "TestExportSavesAsOpenXml"
    Dim exportBook As Workbook

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), "Export should produce a valid workbook."
    Assert.AreEqual CLng(xlOpenXMLWorkbook), CLng(exportBook.FileFormat), _
                    "The export should be written as an open XML workbook whatever the machine default is."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportSavesAsOpenXml", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportContainsNoLeftoverBlankSheet()
    CustomTestSetTitles Assert, "SetupImport", "TestExportContainsNoLeftoverBlankSheet"
    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim expectedCount As Long

    On Error GoTo Fail

    Set exportBook = SharedExportWorkbook()

    Assert.IsTrue Not (exportBook Is Nothing), "Export should produce a valid workbook."

    'The default sheet is "Sheet1" on an English Excel and something else
    'elsewhere, so it is named nowhere here. Every sheet left must be one the
    'service wrote, and the count must match.
    expectedCount = 0
    For Each exportedSheet In exportBook.Worksheets
        Assert.IsTrue IsExpectedExportSheet(exportedSheet.Name), _
                      "Unexpected sheet in the export: " & exportedSheet.Name
        expectedCount = expectedCount + 1
    Next exportedSheet

    Assert.AreEqual CLng(5), CLng(expectedCount), _
                    "The export should carry the five setup sheets and nothing else."
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportContainsNoLeftoverBlankSheet", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("SetupImport")
Public Sub TestExportWithSuppliedWorkbookLeavesItOpen()
    CustomTestSetTitles Assert, "SetupImport", "TestExportWithSuppliedWorkbookLeavesItOpen"
    Dim suppliedBook As Workbook
    Dim suppliedName As String
    Dim svc As SetupImport

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Set suppliedBook = NewWorkbook
    suppliedName = suppliedBook.Name

    Service.DisplayPrompts = False
    Set svc = Service
    svc.Export outwb:=suppliedBook

    'The caller owns a workbook it supplies: nothing is saved, nothing is
    'closed, and no sheet is taken out of it.
    Assert.IsTrue IsWorkbookOpen(suppliedName), "Export should leave a supplied workbook open."
    Assert.AreEqual vbNullString, svc.LastExportFile, "Export should record no file for a supplied workbook."
    Assert.IsTrue ExportWorksheetExists(suppliedBook, DICTIONARY_SHEET_NAME), _
                  "Export should write the setup sheets into the supplied workbook."

    DeleteWorkbook suppliedBook
    Exit Sub

CleanupFailure:
    On Error Resume Next
        DeleteWorkbook suppliedBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportWithSuppliedWorkbookLeavesItOpen", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Tests - the error Check reports
'===============================================================================
'@TestMethod("SetupImport")
Public Sub TestCheckReportsOpenFailureWithProjectError()
    CustomTestSetTitles Assert, "SetupImport", "TestCheckReportsOpenFailureWithProjectError"
    Dim unreadablePath As String
    Dim raisedNumber As Long

    On Error GoTo CleanupFailure

    unreadablePath = WriteUnreadableWorkbook()
    Service.Path = unreadablePath

    On Error Resume Next
        Service.Check True, False, False, False, False
        raisedNumber = Err.Number
        Err.Clear
    On Error GoTo CleanupFailure

    'The value of this one is negative: it fails the day Check starts raising 0
    'or 5 because it read Err after its own cleanup ran.
    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), raisedNumber, _
                    "A file that cannot be opened should be reported as a project error."
    Assert.IsTrue InStr(1, ProgressStub.Value, unreadablePath, vbTextCompare) > 0, _
                  "The message should name the file."

    DeleteFileIfExists unreadablePath
    Exit Sub

CleanupFailure:
    DeleteFileIfExists unreadablePath
    CustomTestLogFailure Assert, "TestCheckReportsOpenFailureWithProjectError", Err.Number, Err.Description
    Err.Clear
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareHostSetupSheets()
    UnprotectIfPossible DICTIONARY_SHEET_NAME
    SetupImportTestFixture.PrepareSetupDictionarySheet DICTIONARY_SHEET_NAME, _
                                                      HOST_DICTIONARY_VARIABLE, _
                                                      "HostSheet", _
                                                      DICTIONARY_HOST_START_ROW, _
                                                      DICTIONARY_HOST_START_COLUMN

    On Error Resume Next
        ThisWorkbook.Names("__ll_exports_total__").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="__ll_exports_total__", RefersTo:="=1"

    UnprotectIfPossible CHOICES_SHEET_NAME
    SetupImportTestFixture.PrepareSetupChoicesSheet CHOICES_SHEET_NAME, _
                                                   CHOICES_HOST_START_ROW, _
                                                   CHOICES_HOST_START_COLUMN

    UnprotectIfPossible EXPORTS_SHEET_NAME
    SetupImportTestFixture.PrepareSetupExportsSheet EXPORTS_SHEET_NAME, _
                                                   HOST_EXPORT_STATUS, _
                                                   HOST_EXPORT_FILE_NAME, _
                                                   HOST_EXPORT_LABEL, _
                                                   EXPORT_HOST_START_ROW, _
                                                   EXPORT_HOST_START_COLUMN

    UnprotectIfPossible ANALYSIS_SHEET_NAME
    SetupImportTestFixture.PrepareSetupAnalysisSheet ANALYSIS_SHEET_NAME, _
                                                    "Host", _
                                                    "Host analysis header"

    UnprotectIfPossible TRANSLATIONS_SHEET_NAME
    SetupImportTestFixture.PrepareSetupTranslationsSheet TRANSLATIONS_SHEET_NAME, _
                                                        TRANSLATIONS_TABLE_NAME, _
                                                        "Host label", _
                                                        HOST_TRANSLATION_VALUE, _
                                                        HOST_TRANSLATION_TAG, _
                                                        TRANSLATION_HOST_START_ROW, _
                                                        TRANSLATION_HOST_START_COLUMN, _
                                                        True

    PrepareRegistryFixture
End Sub

Private Function BuildImportWorkbookFixture() As Workbook
    Dim wb As Workbook

    Set wb = NewWorkbook

    SetupImportTestFixture.PrepareSetupDictionarySheet DICTIONARY_SHEET_NAME, _
                                                      SOURCE_DICTIONARY_VARIABLE, _
                                                      "ImportSheet", _
                                                      SOURCE_START_ROW, _
                                                      SOURCE_START_COLUMN, _
                                                      wb

    SetupImportTestFixture.PrepareSetupChoicesSheet CHOICES_SHEET_NAME, _
                                                   SOURCE_START_ROW, _
                                                   SOURCE_START_COLUMN, _
                                                   wb

    SetupImportTestFixture.PrepareSetupExportsSheet EXPORTS_SHEET_NAME, _
                                                   SOURCE_EXPORT_STATUS, _
                                                   SOURCE_EXPORT_FILE_NAME, _
                                                   SOURCE_EXPORT_LABEL, _
                                                   SOURCE_START_ROW, _
                                                   SOURCE_START_COLUMN, _
                                                   wb

    SetupImportTestFixture.PrepareSetupAnalysisSheet ANALYSIS_SHEET_NAME, _
                                                    "Import", _
                                                    SOURCE_ANALYSIS_HEADER, _
                                                    wb

    SetupImportTestFixture.PrepareSetupTranslationsSheet TRANSLATIONS_SHEET_NAME, _
                                                        TRANSLATIONS_TABLE_NAME, _
                                                        "Import label", _
                                                        SOURCE_TRANSLATION_VALUE, _
                                                        SOURCE_TRANSLATION_TAG, _
                                                        TRANSLATION_SOURCE_START_ROW, _
                                                        TRANSLATION_SOURCE_START_COLUMN, _
                                                        False, _
                                                        wb

    On Error Resume Next
        wb.Names("__ll_exports_total__").Delete
    On Error GoTo 0
    wb.Names.Add Name:="__ll_exports_total__", RefersTo:="=2"

    Set BuildImportWorkbookFixture = wb
End Function

Private Sub UnprotectIfPossible(ByVal sheetName As String)
    If PasswordsHandler Is Nothing Then Exit Sub

    On Error Resume Next
        PasswordsHandler.UnProtect sheetName
    On Error GoTo 0
End Sub

Private Sub PrepareRegistryFixture()
    Dim registrySheet As Worksheet
    Dim dataWksh As Worksheet
    Dim matrix As Variant
    Dim registryRange As Range
    Dim registryTable As ListObject
    Dim store As HiddenNames

    Set dataWksh = EnsureWorksheet(REGISTRY_SOURCE_SHEET)
    dataWksh.Cells.Clear
    dataWksh.Range("A1").Value = SOURCE_TRANSLATION_VALUE
    dataWksh.Range("A2").Value = SOURCE_TRANSLATION_VALUE & " updated"

    On Error Resume Next
        ThisWorkbook.Names(REGISTRY_RANGE_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=REGISTRY_RANGE_NAME, RefersTo:=dataWksh.Range("A1:A2")

    Set registrySheet = EnsureWorksheet(REGISTRY_SHEET_NAME)
    registrySheet.Cells.Clear

    matrix = RowsToMatrix(Array( _
        Array("rngname", "status", "mode"), _
        Array(REGISTRY_RANGE_NAME, "yes", "translate as text")))
    WriteMatrix registrySheet.Cells(1, 1), matrix

    On Error Resume Next
        Do While registrySheet.ListObjects.Count > 0
            registrySheet.ListObjects(1).Delete
        Loop
    On Error GoTo 0

    Set registryRange = registrySheet.Range("A1:C2")
    Set registryTable = registrySheet.ListObjects.Add(xlSrcRange, registryRange, , xlYes)
    registryTable.Name = REGISTRY_TABLE_NAME
    registryTable.TableStyle = ""

    On Error Resume Next
        Set store = HiddenNames.Create(registrySheet)
    On Error GoTo 0
    If Not store Is Nothing Then
        On Error Resume Next
            store.RemoveName REGISTRY_COUNTER_NAME
        On Error GoTo 0
    End If

    On Error Resume Next
        ThisWorkbook.Names(REGISTRY_COUNTER_NAME).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=REGISTRY_COUNTER_NAME, RefersTo:="=0"
End Sub


'@section Import assertion helpers
'===============================================================================
Private Sub AssertImportedDictionary()
    Dim dictSheet As Worksheet
    Dim variableName As String
    Dim exportTotal As Long

    Set dictSheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
    variableName = CStr(dictSheet.Cells(DICTIONARY_HOST_START_ROW + 1, DICTIONARY_HOST_START_COLUMN).Value)

    Assert.AreEqual SOURCE_DICTIONARY_VARIABLE, variableName, "Dictionary import should replace the variable name."

    exportTotal = HostExportTotal()
    Assert.AreEqual CLng(1), exportTotal, "Dictionary import should keep the export counter unchanged."
End Sub

Private Sub AssertImportedExports()
    Dim exportSheet As Worksheet
    Dim lo As ListObject
    Dim statusIdx As Long
    Dim fileIdx As Long

    Set exportSheet = ThisWorkbook.Worksheets(EXPORTS_SHEET_NAME)
    Set lo = exportSheet.ListObjects(1)

    statusIdx = lo.ListColumns("status").Index
    fileIdx = lo.ListColumns("file name").Index

    Assert.AreEqual SOURCE_EXPORT_STATUS, CStr(lo.DataBodyRange.Cells(1, statusIdx).Value), _
                    "Exports import should replace the status field."
    Assert.AreEqual SOURCE_EXPORT_FILE_NAME, CStr(lo.DataBodyRange.Cells(1, fileIdx).Value), _
                    "Exports import should replace the file name."
End Sub

Private Sub AssertImportedAnalysis()
    Dim analysisSheet As Worksheet
    Dim summaryTable As ListObject

    Set analysisSheet = ThisWorkbook.Worksheets(ANALYSIS_SHEET_NAME)

    Set summaryTable = analysisSheet.ListObjects("Tab_global_summary")
    Assert.AreEqual "Import global section", _
                    CStr(summaryTable.DataBodyRange.Cells(1, 1).Value), _
                    "Analysis import should copy table rows."
End Sub

Private Sub AssertImportedTranslations()
    Dim translationSheet As Worksheet
    Dim lo As ListObject
    Dim labelIdx As Long
    Dim englishIdx As Long
    Dim firstTag As String
    Dim secondTag As String

    Set translationSheet = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
    Set lo = translationSheet.ListObjects(TRANSLATIONS_TABLE_NAME)

    labelIdx = lo.ListColumns("Lang1").Index
    Assert.AreEqual "Import Label", _
                    CStr(lo.DataBodyRange.Cells(1, labelIdx).Value), _
                    "Translations import should keep existing lang1 values."

    'Ensure headers from the source workbook are preserved.
    Assert.AreEqual "English", lo.ListColumns("English").Name, _
                    "Translations import should keep existing headers."

    Assert.AreEqual CLng(1), CLng(lo.ListRows.Count), _
                    "Translations import should rebuild the table based on imported data."

    englishIdx = lo.ListColumns("English").Index
    Assert.AreEqual SOURCE_TRANSLATION_VALUE, _
                    CStr(lo.DataBodyRange.Cells(1, englishIdx).Value), _
                    "Translations import should copy the English values from the source table."

    'The host tag described a row the import wrote over; every imported row
    'now carries the imported marker, for the next update to review.
    firstTag = CStr(translationSheet.Cells(TRANSLATION_HOST_START_ROW + 1, TRANSLATION_HOST_START_COLUMN - 1).Value)
    Assert.AreEqual "__imported____0", firstTag, _
                    "Translations import should mark every imported row for review."

    Assert.AreEqual CLng(0), RegistryCounterValue(), _
                    "Translations registry counter should remain unchanged after import."
End Sub

Private Sub AssertTranslationUnchanged()
    Dim translationSheet As Worksheet
    Dim lo As ListObject
    Dim columnIdx As Long
    Dim tagValue As String

    Set translationSheet = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
    Set lo = translationSheet.ListObjects(TRANSLATIONS_TABLE_NAME)

    columnIdx = lo.ListColumns("English").Index
    Assert.AreEqual HOST_TRANSLATION_VALUE, _
                    CStr(lo.DataBodyRange.Cells(1, columnIdx).Value), _
                    "Translations import should not alter values when the source sheet is missing."

    tagValue = CStr(translationSheet.Cells(TRANSLATION_HOST_START_ROW + 1, TRANSLATION_HOST_START_COLUMN - 1).Value)
    Assert.AreEqual HOST_TRANSLATION_TAG, tagValue, _
                    "Translations import should keep existing tags when the source sheet is missing."

    Assert.AreEqual CLng(0), RegistryCounterValue(), _
                    "Registry counter should remain unchanged when no translation import occurs."
End Sub


'@section Export helpers
'===============================================================================

'@description Execute the full export workflow and open the resulting workbook for verification.
'@param exportFilePath ByRef String receiving the path to the exported file.
'@return Workbook opened from the exported file, or Nothing if export did not produce a file.
Private Function PerformExportAndOpen(ByRef exportFilePath As String) As Workbook
    Dim exportFolder As String
    Dim svc As SetupImport

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupExportTests")

    Service.DisplayPrompts = False
    Service.SetExportFolder exportFolder

    Set svc = Service
    svc.Export

    'The file name carries a time stamp, so the path is read back from the
    'service rather than worked out again here.
    exportFilePath = svc.LastExportFile
    If LenB(exportFilePath) = 0 Then Exit Function
    If LenB(Dir$(exportFilePath)) = 0 Then Exit Function

    Set PerformExportAndOpen = Workbooks.Open(exportFilePath)
End Function

'@description Answer the one export workbook the read-only export tests share.
'@details Fifteen tests read the workbook Export writes and change nothing in
'   it. Each of them used to build the five host sheets, run a full export,
'   save, open, close and delete a real file, and together they were the bulk
'   of this module's wall clock. The export runs once here, on the first call,
'   and every later call answers the same open workbook. ModuleCleanup closes
'   it and deletes the file.
'@return Workbook the shared export workbook, already open.
Private Function SharedExportWorkbook() As Workbook
    If Not SharedExportBook Is Nothing Then
        Set SharedExportWorkbook = SharedExportBook
        Exit Function
    End If

    PrepareHostSetupSheets
    Set SharedExportBook = PerformExportAndOpen(SharedExportPath)
    Set SharedExportWorkbook = SharedExportBook
End Function

'@description Answer the file path the shared export was written to.
'@return String path of the shared export file.
Private Function SharedExportFilePath() As String
    If SharedExportBook Is Nothing Then
        Set SharedExportBook = SharedExportWorkbook()
    End If
    SharedExportFilePath = SharedExportPath
End Function

'@description Close the shared export workbook and delete its file.
Private Sub ReleaseSharedExport()
    On Error Resume Next
        If Not SharedExportBook Is Nothing Then SharedExportBook.Close SaveChanges:=False
    On Error GoTo 0
    Set SharedExportBook = Nothing

    If Not KeepExportArtifacts Then
        DeleteFileIfExists SharedExportPath
    End If
    SharedExportPath = vbNullString
End Sub

'@description Close an export workbook and delete the file if artifacts are not kept.
'@param exportBook ByRef Workbook reference to close and release.
'@param exportFilePath String path to the exported file to delete.
Private Sub CleanupExportResult(ByRef exportBook As Workbook, ByVal exportFilePath As String)
    On Error Resume Next
        If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    On Error GoTo 0
    Set exportBook = Nothing
    If Not KeepExportArtifacts Then
        DeleteFileIfExists exportFilePath
    End If
End Sub

'@description Check if a worksheet exists in a given workbook.
'@param wb Workbook to search.
'@param sheetName String name of the worksheet.
'@return Boolean True when the worksheet is present.
Private Function ExportWorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet

    If wb Is Nothing Then Exit Function

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    ExportWorksheetExists = Not (sh Is Nothing)
End Function

'@description Host workbook name with whatever follows the last dot removed.
'@return String the base name the export file is built from.
Private Function HostBaseName() As String
    Dim baseName As String
    Dim dotPosition As Long

    baseName = ThisWorkbook.Name
    dotPosition = InStrRev(baseName, ".")
    If dotPosition > 1 Then baseName = Left$(baseName, dotPosition - 1)

    HostBaseName = baseName
End Function

'@description Count the columns of a table carrying a header, ignoring case.
'@param lo ListObject to walk.
'@param headerName String header to look for.
'@return Long number of columns holding that header.
Private Function HeaderCount(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim headerCell As Range
    Dim total As Long

    If lo Is Nothing Then Exit Function

    For Each headerCell In lo.HeaderRowRange.Cells
        If StrComp(Trim$(CStr(headerCell.Value)), Trim$(headerName), vbTextCompare) = 0 Then
            total = total + 1
        End If
    Next headerCell

    HeaderCount = total
End Function

'@description Hold the run until the clock second changes.
'@details The export file name is stamped to the second, so two exports inside
'   one second would resolve to one path. This is what makes the uniqueness
'   test say something.
Private Sub WaitForNextSecond()
    Dim startTime As Double

    startTime = Timer
    Do While Timer >= startTime And Timer - startTime < 1.1
    Loop
End Sub

'@description Find the column of an exported sheet carrying a given header.
'@param exportedSheet Worksheet to read, headers on row 1.
'@param headerName String header to look for, matched without regard to case.
'@return Long the column number, 0 when the header is absent.
Private Function ExportedHeaderColumn(ByVal exportedSheet As Worksheet, _
                                      ByVal headerName As String) As Long
    Dim lastColumn As Long
    Dim colIdx As Long

    If exportedSheet Is Nothing Then Exit Function

    lastColumn = exportedSheet.Cells(1, exportedSheet.Columns.Count).End(xlToLeft).Column

    For colIdx = 1 To lastColumn
        If StrComp(Trim$(CStr(exportedSheet.Cells(1, colIdx).Value)), _
                   Trim$(headerName), vbTextCompare) = 0 Then
            ExportedHeaderColumn = colIdx
            Exit Function
        End If
    Next colIdx
End Function

'@description Answer whether a sheet name is one the export puts in the workbook.
'@param sheetName String name to test.
'@return Boolean True when the service wrote that sheet.
Private Function IsExpectedExportSheet(ByVal sheetName As String) As Boolean
    Select Case LCase$(Trim$(sheetName))
        Case LCase$(DICTIONARY_SHEET_NAME), LCase$(CHOICES_SHEET_NAME), _
             LCase$(EXPORTS_SHEET_NAME), LCase$(ANALYSIS_SHEET_NAME), _
             LCase$(TRANSLATIONS_SHEET_NAME), "__formatter"
            IsExpectedExportSheet = True
    End Select
End Function


'@section Utility helpers
'===============================================================================
Private Function HostExportTotal() As Long
    Dim definition As Name
    Dim evaluated As String

    On Error Resume Next
        Set definition = ThisWorkbook.Names("__ll_exports_total__")
    On Error GoTo 0

    If definition Is Nothing Then Exit Function

    evaluated = Replace(definition.Value, "=", vbNullString)
    If LenB(Trim$(evaluated)) > 0 Then
        HostExportTotal = CLng(Trim$(evaluated))
    End If
End Function

Private Function RegistryCounterValue() As Long
    Dim registrySheet As Worksheet
    Dim store As HiddenNames

    On Error Resume Next
        Set registrySheet = ThisWorkbook.Worksheets(REGISTRY_SHEET_NAME)
    On Error GoTo 0
    If registrySheet Is Nothing Then Exit Function

    On Error Resume Next
        Set store = HiddenNames.Create(registrySheet)
    On Error GoTo 0
    If store Is Nothing Then Exit Function

    RegistryCounterValue = store.ValueAsLong(REGISTRY_COUNTER_NAME, 0)
End Function

Private Sub EnsurePasswordsFixture()
    Dim passwordSheet As Worksheet

    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, ThisWorkbook
    Set passwordSheet = ThisWorkbook.Worksheets(PASSWORD_SHEET)
    Set PasswordsHandler = Passwords.Create(passwordSheet)
End Sub

Private Sub PrepareComment(ByVal targetSheet As Worksheet)
    On Error Resume Next
        targetSheet.Cells(1, 1).ClearComments
        targetSheet.Cells(1, 1).ClearCommentsThreaded
    On Error GoTo 0

    targetSheet.Cells(1, 1).Value = "Sample"
    targetSheet.Cells(1, 1).AddComment "Temporary note"
End Sub

Private Function SheetsListOf(ParamArray sheetNames() As Variant) As BetterArray
    Dim list As BetterArray
    Dim idx As Long

    Set list = New BetterArray
    list.LowerBound = 1

    For idx = LBound(sheetNames) To UBound(sheetNames)
        list.Push CStr(sheetNames(idx))
    Next idx

    Set SheetsListOf = list
End Function

Private Function BuildMissingSetupPath() As String
    Dim baseFolder As String

    baseFolder = ThisWorkbook.Path
    If LenB(baseFolder) = 0 Then baseFolder = CurDir$

    BuildMissingSetupPath = baseFolder & Application.PathSeparator & "missing_setup_source.xlsx"
End Function

Private Function IsWorkbookOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function FileNameFromPath(ByVal filePath As String) As String
    Dim separatorPos As Long

    separatorPos = InStrRev(filePath, Application.PathSeparator)
    If separatorPos = 0 Then
        FileNameFromPath = filePath
    Else
        FileNameFromPath = Mid$(filePath, separatorPos + 1)
    End If
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    If LenB(filePath) = 0 Then Exit Sub
    If LenB(Dir$(filePath)) = 0 Then Exit Sub

    On Error Resume Next
        Kill filePath
    On Error GoTo 0
End Sub


'@section Helpers for the session 25 regression tests
'===============================================================================
'@description Give the host Choices sheet a table.
'@details The import path matches source and host tables by name, and
'   PrepareImport renames two columns of the first table on the sheet. Neither
'   reaches a Choices sheet that carries no table.
Private Sub PrepareHostChoicesTable()
    UnprotectIfPossible CHOICES_SHEET_NAME
    SetupImportTestFixture.PrepareSetupChoicesSheet CHOICES_SHEET_NAME, _
                                                    CHOICES_HOST_START_ROW, _
                                                    CHOICES_HOST_START_COLUMN, _
                                                    tableName:=CHOICES_TABLE_NAME
End Sub

'@description Resolve the first table of the host Choices sheet.
'@return ListObject, or Nothing when the sheet or the table is missing.
Private Function HostChoicesTable() As ListObject
    On Error Resume Next
        Set HostChoicesTable = ThisWorkbook.Worksheets(CHOICES_SHEET_NAME).ListObjects(1)
        Err.Clear
    On Error GoTo 0
End Function

'@description Resolve the first table of the host Exports sheet.
'@return ListObject, or Nothing when the sheet or the table is missing.
Private Function HostExportsTable() As ListObject
    On Error Resume Next
        Set HostExportsTable = ThisWorkbook.Worksheets(EXPORTS_SHEET_NAME).ListObjects(1)
        Err.Clear
    On Error GoTo 0
End Function

'@description Build a source setup workbook, save it, and return its path.
'@param fileTag String name fragment for the saved file.
'@param extraExportRows Long extra rows appended to the source exports table.
'@param exportsOnly Boolean True to leave the Choices sheet out.
'@return String path to the saved workbook.
Private Function BuildSetupSourceFile(ByVal fileTag As String, _
                                      Optional ByVal extraExportRows As Long = 0, _
                                      Optional ByVal exportsOnly As Boolean = False) As String
    Dim wb As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim exportsSheet As Worksheet
    Dim idx As Long

    Set wb = NewWorkbook

    SetupImportTestFixture.PrepareSetupDictionarySheet DICTIONARY_SHEET_NAME, _
                                                       SOURCE_DICTIONARY_VARIABLE, _
                                                       "ImportSheet", _
                                                       SOURCE_START_ROW, _
                                                       SOURCE_START_COLUMN, _
                                                       wb

    SetupImportTestFixture.PrepareSetupExportsSheet EXPORTS_SHEET_NAME, _
                                                    SOURCE_EXPORT_STATUS, _
                                                    SOURCE_EXPORT_FILE_NAME, _
                                                    SOURCE_EXPORT_LABEL, _
                                                    SOURCE_START_ROW, _
                                                    SOURCE_START_COLUMN, _
                                                    wb

    If Not exportsOnly Then
        SetupImportTestFixture.PrepareSetupChoicesSheet CHOICES_SHEET_NAME, _
                                                        SOURCE_START_ROW, _
                                                        SOURCE_START_COLUMN, _
                                                        wb, _
                                                        CHOICES_TABLE_NAME
    End If

    If extraExportRows > 0 Then
        Set exportsSheet = wb.Worksheets(EXPORTS_SHEET_NAME)
        For idx = 1 To extraExportRows
            AppendExportRow exportsSheet.ListObjects(1), idx + 1
        Next idx
    End If

    exportFolder = BuildTempFolder(ThisWorkbook, "SetupImportTests")
    workbookPath = BuildWorkbookPath(exportFolder, fileTag, ".xlsx")
    DeleteFileIfExists workbookPath

    wb.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False

    BuildSetupSourceFile = workbookPath
End Function

'@description Copy the last row of an exports table and give the copy a number.
'@param lo ListObject holding the export rows.
'@param exportNumber Long number written into the new row.
Private Sub AppendExportRow(ByVal lo As ListObject, ByVal exportNumber As Long)
    Dim newRow As ListRow
    Dim previousValues As Variant
    Dim colIdx As Long

    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    previousValues = lo.ListRows(lo.ListRows.Count).Range.Value
    Set newRow = lo.ListRows.Add

    For colIdx = 1 To lo.ListColumns.Count
        newRow.Range.Cells(1, colIdx).Value = previousValues(1, colIdx)
    Next colIdx

    newRow.Range.Cells(1, lo.ListColumns("export number").Index).Value = exportNumber
    newRow.Range.Cells(1, lo.ListColumns("file name").Index).Value = "extra_" & CStr(exportNumber) & ".xlsx"
End Sub

'@description Lock a host sheet with a key the Passwords handler does not hold.
'@param sheetName String name of the sheet to lock.
Private Sub LockSheetWithForeignKey(ByVal sheetName As String)
    Dim targetSheet As Worksheet

    On Error Resume Next
        'Take the handler's own protection off first. Calling Unprotect with no
        'password on a protected sheet opens a password dialog, and a dialog
        'stops the whole run.
        PasswordsHandler.UnProtect sheetName
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
        If Not targetSheet Is Nothing Then
            targetSheet.Protect Password:=FOREIGN_PASSWORD
        End If
        Err.Clear
    On Error GoTo 0
End Sub

'@description Take the foreign key protection back off a host sheet.
'@param sheetName String name of the sheet to unlock.
Private Sub UnlockForeignKeySheet(ByVal sheetName As String)
    Dim targetSheet As Worksheet

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
        If Not targetSheet Is Nothing Then
            targetSheet.Unprotect FOREIGN_PASSWORD
        End If
        Err.Clear
    On Error GoTo 0
End Sub

'@description Write a text file carrying a workbook name, so opening it fails.
'@return String path to the file.
Private Function WriteUnreadableWorkbook() As String
    Dim folderPath As String
    Dim filePath As String
    Dim fileNumber As Integer

    folderPath = BuildTempFolder(ThisWorkbook, "SetupImportTests")
    filePath = BuildWorkbookPath(folderPath, "setup_unreadable", ".xlsx")
    DeleteFileIfExists filePath

    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, "this file carries a workbook name and no workbook"
    Close #fileNumber

    WriteUnreadableWorkbook = filePath
End Function

'@description Report an unexpected error and drop the file the test wrote.
'@param routineName String name of the test.
'@param filePath String file to remove, empty when there is none.
Private Sub LogUnexpected(ByVal routineName As String, ByVal filePath As String)
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
        DeleteFileIfExists filePath
    On Error GoTo 0

    If errNumber <> 0 Then
        CustomTestLogFailure Assert, routineName, errNumber, errDescription
        Err.Clear
    End If
End Sub
