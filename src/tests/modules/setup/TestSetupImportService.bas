Attribute VB_Name = "TestSetupImportService"
Option Explicit


'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the improved setup import service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Service As ISetupImportService
Private ProgressStub As ProgressDisplayStub
Private PasswordsHandler As IPasswords

Private Const PASSWORD_SHEET As String = "TST_SetupImport_Passwords"
Private Const CLEAN_TARGET_SHEET As String = "TST_SetupImport_Clean"
Private Const DICTIONARY_SHEET_NAME As String = "Dictionary"
Private Const EXPORTS_SHEET_NAME As String = "Exports"
Private Const ANALYSIS_SHEET_NAME As String = "Analysis"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
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

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupImportService"
    KeepExportArtifacts = False
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    Set ProgressStub = New ProgressDisplayStub
    ProgressStub.Caption = vbNullString
    ProgressStub.Value = vbNullString
    Set Service = New SetupImportService
    Service.Path = ThisWorkbook.FullName
    Set Service.ProgressObject = ProgressStub
    EnsurePasswordsFixture
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Service = Nothing
    Set ProgressStub = Nothing
    Set PasswordsHandler = Nothing
    TestHelpers.DeleteWorksheet CLEAN_TARGET_SHEET
    TestHelpers.DeleteWorksheet PASSWORD_SHEET
    TestHelpers.DeleteWorksheet REGISTRY_SHEET_NAME
    TestHelpers.DeleteWorksheet REGISTRY_SOURCE_SHEET
    TestHelpers.DeleteWorksheet CHOICES_SHEET_NAME
    TestHelpers.DeleteWorksheet DICTIONARY_SHEET_NAME
    TestHelpers.DeleteWorksheet EXPORTS_SHEET_NAME
    TestHelpers.DeleteWorksheet ANALYSIS_SHEET_NAME
    TestHelpers.DeleteWorksheet TRANSLATIONS_SHEET_NAME
    On Error Resume Next
        ThisWorkbook.Names(REGISTRY_RANGE_NAME).Delete
    On Error GoTo 0
End Sub


'@section Tests - Check and Clean
'===============================================================================
'@TestMethod("SetupImportService")
Public Sub TestCheckRaisesWhenNoSelection()
    CustomTestSetTitles Assert, "SetupImportService", "TestCheckRaisesWhenNoSelection"
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

'@TestMethod("SetupImportService")
Public Sub TestCheckRaisesWhenFileMissing()
    CustomTestSetTitles Assert, "SetupImportService", "TestCheckRaisesWhenFileMissing"
    Dim missingPath As String

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
End Sub

'@TestMethod("SetupImportService")
Public Sub TestCleanRemovesWorksheetComments()
    CustomTestSetTitles Assert, "SetupImportService", "TestCleanRemovesWorksheetComments"
    Dim targetSheet As Worksheet
    Dim sheetsList As BetterArray

    Set targetSheet = TestHelpers.EnsureWorksheet(CLEAN_TARGET_SHEET)
    PrepareComment targetSheet

    Set sheetsList = SheetsListOf(CLEAN_TARGET_SHEET)
    Service.Clean PasswordsHandler, sheetsList

    Assert.IsTrue targetSheet.Cells(1, 1).Comment Is Nothing, "Clean should remove classic comments."
End Sub


'@section Tests - Import
'===============================================================================
'@TestMethod("SetupImportService")
Public Sub TestImportClosesWorkbookAfterRun()
    CustomTestSetTitles Assert, "SetupImportService", "TestImportClosesWorkbookAfterRun"
    Dim tempBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim workbookName As String

    Set tempBook = TestHelpers.NewWorkbook
    tempBook.Worksheets(1).Name = "TempData"

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupImportTests")
    workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "setup_import_source", ".xlsx")
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
End Sub

'@TestMethod("SetupImportService")
Public Sub TestImportFromWorkbookUsingDomainClasses()
    CustomTestSetTitles Assert, "SetupImportService", "TestImportFromWorkbookUsingDomainClasses"
    Dim sourceBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim workbookName As String
    Dim sheetsList As BetterArray

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    Set sourceBook = BuildImportWorkbookFixture()
    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupImportDomainTests")
    workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "setup_import_domain", ".xlsx")

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

'@TestMethod("SetupImportService")
Public Sub TestImportFromWorkbookSkipsMissingSheets()
    CustomTestSetTitles Assert, "SetupImportService", "TestImportFromWorkbookSkipsMissingSheets"
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

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupImportDomainTests")
    workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "setup_import_missing", ".xlsx")

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
'@TestMethod("SetupImportService")
Public Sub TestExportAbortsWhenFolderSelectionCancelled()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportAbortsWhenFolderSelectionCancelled"
    Dim initialWorkbookCount As Long
    Dim svc As ISetupImportService

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
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportCreatesWorkbookInProvidedFolder()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportCreatesWorkbookInProvidedFolder"
    Dim exportFolder As String
    Dim expectedFilePath As String
    Dim svc As ISetupImportService
    Dim initialWorkbookCount As Long

    PrepareHostSetupSheets

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupExportTests")
    expectedFilePath = exportFolder & Application.PathSeparator & Replace(ThisWorkbook.Name, ".xlsb", "") & "_export_" & Format$(Now(), "yyyymmdd") & ".xlsx"
    DeleteFileIfExists expectedFilePath

    Service.DisplayPrompts = False
    Service.SetExportFolder exportFolder

    initialWorkbookCount = Application.Workbooks.Count
    Set svc = Service
    svc.Export

    Assert.IsTrue LenB(Dir$(expectedFilePath)) > 0, "Export should write the workbook to the configured folder."
    Assert.AreEqual initialWorkbookCount, Application.Workbooks.Count, "Export should close the temporary export workbook."
    Assert.AreEqual expectedFilePath, svc.LastExportFile, "Export should expose the saved file path."

    Dim exportBook As Workbook
    Dim translationSheet As Worksheet
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo ExportVerificationFailed
        Set exportBook = Workbooks.Open(expectedFilePath)
        Set translationSheet = exportBook.Worksheets(TRANSLATIONS_SHEET_NAME)
        Assert.AreEqual "lang1", LCase$(CStr(translationSheet.Cells(1, 2).Value)), _
                        "Translations export should include the label column header starting on the second column."
        Assert.AreEqual "english", LCase$(CStr(translationSheet.Cells(1, 3).Value)), _
                        "Translations export should include the English column header offset by one column."
        Assert.AreEqual HOST_TRANSLATION_VALUE, CStr(translationSheet.Cells(2, 3).Value), _
                        "Translations export should retain existing translations when shifted to the second column."
        exportBook.Close SaveChanges:=False
    On Error GoTo 0

    If Not KeepExportArtifacts Then
        DeleteFileIfExists expectedFilePath
    End If
    Exit Sub

ExportVerificationFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    On Error GoTo 0
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

'@TestMethod("SetupImportService")
Public Sub TestExportContainsDictionarySheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportContainsDictionarySheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportContainsDictionarySheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportContainsChoicesSheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportContainsChoicesSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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
    Assert.AreEqual "Choice A", CStr(exportedSheet.Cells(2, 3).Value), _
                    "Choices export should include the label column data."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportContainsChoicesSheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportContainsExportSpecsSheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportContainsExportSpecsSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet
    Dim lastColumn As Long
    Dim colIdx As Long
    Dim statusColumn As Long
    Dim fileNameColumn As Long
    Dim headerValue As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportContainsExportSpecsSheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportContainsAnalysisSheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportContainsAnalysisSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet
    Dim summaryTable As ListObject

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportContainsAnalysisSheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportContainsMultipleAnalysisTables()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportContainsMultipleAnalysisTables"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet
    Dim univariateTable As ListObject
    Dim timeseriesTable As ListObject

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportContainsMultipleAnalysisTables", errNumber, errDescription
        Err.Clear
    End If
End Sub


'@section Tests - Export structural checks
'===============================================================================

'@TestMethod("SetupImportService")
Public Sub TestExportRemovesDefaultWorksheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportRemovesDefaultWorksheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    'The default Sheet1 from Workbooks.Add should have been removed
    Assert.IsFalse ExportWorksheetExists(exportBook, "Sheet1"), _
                   "Export should remove the default Sheet1 worksheet."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportRemovesDefaultWorksheet", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportMainSheetsAreVisible()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportMainSheetsAreVisible"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportMainSheetsAreVisible", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportClosesWorkbookAfterCompletion()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportClosesWorkbookAfterCompletion"
    Dim exportFilePath As String
    Dim initialCount As Long
    Dim svc As ISetupImportService

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    exportFilePath = BuildExportFilePath()
    DeleteFileIfExists exportFilePath

    Service.DisplayPrompts = False
    Service.SetExportFolder TestHelpers.BuildTempFolder(ThisWorkbook, "SetupExportTests")

    initialCount = Application.Workbooks.Count
    Set svc = Service
    svc.Export

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

'@TestMethod("SetupImportService")
Public Sub TestExportSkipsMissingHostDictionarySheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportSkipsMissingHostDictionarySheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Dictionary sheet from the host before exporting
    TestHelpers.DeleteWorksheet DICTIONARY_SHEET_NAME

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

'@TestMethod("SetupImportService")
Public Sub TestExportSkipsMissingHostAnalysisSheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportSkipsMissingHostAnalysisSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Analysis sheet from the host before exporting
    TestHelpers.DeleteWorksheet ANALYSIS_SHEET_NAME

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

'@TestMethod("SetupImportService")
Public Sub TestExportSkipsMissingHostChoicesSheet()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportSkipsMissingHostChoicesSheet"
    Dim exportBook As Workbook
    Dim exportFilePath As String

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets

    'Remove the Choices sheet from the host before exporting
    TestHelpers.DeleteWorksheet CHOICES_SHEET_NAME

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

'@TestMethod("SetupImportService")
Public Sub TestExportDictionaryContainsListObject()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportDictionaryContainsListObject"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(DICTIONARY_SHEET_NAME)

    'LLdictionary.Export creates a ListObject in the exported sheet
    Assert.IsTrue exportedSheet.ListObjects.Count > 0, _
                  "Dictionary export should contain a ListObject wrapping the data."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportDictionaryContainsListObject", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportDictionaryPreservesMultipleHeaders()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportDictionaryPreservesMultipleHeaders"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(DICTIONARY_SHEET_NAME)

    'Verify several key dictionary headers are present in the export
    Assert.AreEqual "main label", LCase$(CStr(exportedSheet.Cells(1, 2).Value)), _
                    "Dictionary export should include Main Label as the second header."
    Assert.AreEqual "status", LCase$(CStr(exportedSheet.Cells(1, 11).Value)), _
                    "Dictionary export should include Status header."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportDictionaryPreservesMultipleHeaders", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportChoicesPreservesAllDataRows()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportChoicesPreservesAllDataRows"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet
    Dim lastDataRow As Long

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

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

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportChoicesPreservesAllDataRows", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportExportsHiddenNames()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportExportsHiddenNames"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedName As Name

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    'HiddenNames.ExportNamesToWorkbook should copy named ranges to the export workbook
    'The host has __ll_exports_total__ set to 1
    On Error Resume Next
        Set exportedName = exportBook.Names("__ll_exports_total__")
    On Error GoTo 0

    Assert.IsTrue Not (exportedName Is Nothing), _
                  "Export should include the __ll_exports_total__ hidden name in the workbook."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportExportsHiddenNames", errNumber, errDescription
        Err.Clear
    End If
End Sub

'@TestMethod("SetupImportService")
Public Sub TestExportTranslationsStartsAtColumnTwo()
    CustomTestSetTitles Assert, "SetupImportService", "TestExportTranslationsStartsAtColumnTwo"
    Dim exportBook As Workbook
    Dim exportFilePath As String
    Dim exportedSheet As Worksheet

    On Error GoTo CleanupFailure

    PrepareHostSetupSheets
    Set exportBook = PerformExportAndOpen(exportFilePath)

    Assert.IsTrue Not (exportBook Is Nothing), _
                  "Export should produce a valid workbook."

    Set exportedSheet = exportBook.Worksheets(TRANSLATIONS_SHEET_NAME)

    'Translations are exported starting at column 2 to preserve the tag offset
    Assert.AreEqual vbNullString, Trim$(CStr(exportedSheet.Cells(1, 1).Value)), _
                    "Translations export column 1 should be empty (tag offset preserved)."
    Assert.AreEqual "lang1", LCase$(CStr(exportedSheet.Cells(1, 2).Value)), _
                    "Translations export should start headers at column 2."

    CleanupExportResult exportBook, exportFilePath
    Exit Sub

CleanupFailure:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    CleanupExportResult exportBook, exportFilePath
    If errNumber <> 0 Then
        CustomTestLogFailure Assert, "TestExportTranslationsStartsAtColumnTwo", errNumber, errDescription
        Err.Clear
    End If
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

    Set wb = TestHelpers.NewWorkbook

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
    Dim dataSheet As Worksheet
    Dim matrix As Variant
    Dim registryRange As Range
    Dim registryTable As ListObject
    Dim store As IHiddenNames

    Set dataSheet = TestHelpers.EnsureWorksheet(REGISTRY_SOURCE_SHEET)
    dataSheet.Cells.Clear
    dataSheet.Range("A1").Value = SOURCE_TRANSLATION_VALUE
    dataSheet.Range("A2").Value = SOURCE_TRANSLATION_VALUE & " updated"

    On Error Resume Next
        ThisWorkbook.Names(REGISTRY_RANGE_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=REGISTRY_RANGE_NAME, RefersTo:=dataSheet.Range("A1:A2")

    Set registrySheet = TestHelpers.EnsureWorksheet(REGISTRY_SHEET_NAME)
    registrySheet.Cells.Clear

    matrix = TestHelpers.RowsToMatrix(Array( _
        Array("rngname", "status", "mode"), _
        Array(REGISTRY_RANGE_NAME, "yes", "translate as text")))
    TestHelpers.WriteMatrix registrySheet.Cells(1, 1), matrix

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

    firstTag = CStr(translationSheet.Cells(TRANSLATION_HOST_START_ROW + 1, TRANSLATION_HOST_START_COLUMN - 1).Value)
    Assert.AreEqual HOST_TRANSLATION_TAG, firstTag, _
                    "Translations import should leave existing tags untouched."

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
    Dim svc As ISetupImportService

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupExportTests")
    exportFilePath = BuildExportFilePath(exportFolder)
    DeleteFileIfExists exportFilePath

    Service.DisplayPrompts = False
    Service.SetExportFolder exportFolder

    Set svc = Service
    svc.Export

    If LenB(Dir$(exportFilePath)) = 0 Then Exit Function

    Set PerformExportAndOpen = Workbooks.Open(exportFilePath)
End Function

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

'@description Build the expected export file path for the given folder.
'@param exportFolder Optional String folder path. When empty, uses the standard test folder.
'@return String fully qualified path to the expected export file.
Private Function BuildExportFilePath(Optional ByVal exportFolder As String = vbNullString) As String
    If LenB(exportFolder) = 0 Then
        exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupExportTests")
    End If

    BuildExportFilePath = exportFolder & Application.PathSeparator & _
                          Replace(ThisWorkbook.Name, ".xlsb", "") & _
                          "_export_" & Format$(Now(), "yyyymmdd") & ".xlsx"
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
    Dim store As IHiddenNames

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
    If LenB(Dir$(filePath)) = 0 Then Exit Sub

    On Error Resume Next
        Kill filePath
    On Error GoTo 0
End Sub
