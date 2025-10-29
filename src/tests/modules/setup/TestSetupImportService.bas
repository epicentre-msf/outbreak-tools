Attribute VB_Name = "TestSetupImportService"
Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


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
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const REGISTRY_SOURCE_SHEET As String = "TST_SetupImport_RegistrySource"
Private Const REGISTRY_TABLE_NAME As String = "TST_Registry"
Private Const REGISTRY_RANGE_NAME As String = "RNG_HostMessages"
Private Const REGISTRY_COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const DICTIONARY_HEADERS_DEFINITION As String = "variable name|main label|dev comments|editable label|sub label|note|sheet name|sheet type|main section|sub section|status|register book|personal identifier|variable type|variable format|control|control details|unique|min|max|alert|message|formatting condition|formatting values|lock cells"
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
Private Const TRANSLATION_HOST_START_ROW As Long = 5
Private Const TRANSLATION_HOST_START_COLUMN As Long = 2
Private Const SOURCE_START_ROW As Long = 1
Private Const SOURCE_START_COLUMN As Long = 1
Private Const TRANSLATION_SOURCE_START_ROW As Long = 1
Private Const TRANSLATION_SOURCE_START_COLUMN As Long = 1


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupImportService"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set ProgressStub = New ProgressDisplayStub
    Set Service = New SetupImportService
    Service.Path = ThisWorkbook.FullName
    Set Service.ProgressObject = ProgressStub
    EnsurePasswordsFixture
End Sub

'@TestCleanup
Private Sub TestCleanup()
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
    On Error Resume Next
        ThisWorkbook.Names(REGISTRY_RANGE_NAME).Delete
    On Error GoTo 0
End Sub


'@section Tests
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
                    ProgressStub.Caption, "Expected message to be surfaced through the progress display."
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
    Assert.IsTrue InStr(1, ProgressStub.Caption, missingPath, vbTextCompare) > 0, _
                   "Progress display should include the missing path."
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
    On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    DeleteFileIfExists workbookPath
    Err.Raise Err.Number, Err.Source, Err.Description
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
    On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    DeleteFileIfExists workbookPath
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareHostSetupSheets()
    Dim dictSheet As Worksheet
    Dim exportSheet As Worksheet
    Dim analysisSheet As Worksheet
    Dim translationSheet As Worksheet

    Set dictSheet = TestHelpers.EnsureWorksheet(DICTIONARY_SHEET_NAME)
    PopulateDictionarySheet dictSheet, HOST_DICTIONARY_VARIABLE, "HostSheet", DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN

    On Error Resume Next
        ThisWorkbook.Names("__ll_exports_total__").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="__ll_exports_total__", RefersTo:="=1"

    Set exportSheet = TestHelpers.EnsureWorksheet(EXPORTS_SHEET_NAME)
    PopulateExportsSheet exportSheet, HOST_EXPORT_STATUS, HOST_EXPORT_FILE_NAME, HOST_EXPORT_LABEL, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN

    Set analysisSheet = TestHelpers.EnsureWorksheet(ANALYSIS_SHEET_NAME)
    PopulateAnalysisSheet analysisSheet, "Host", "Host analysis header"

    Set translationSheet = TestHelpers.EnsureWorksheet(TRANSLATIONS_SHEET_NAME)
    PopulateTranslationsSheet translationSheet, "Host label", HOST_TRANSLATION_VALUE, HOST_TRANSLATION_TAG, _
                             TRANSLATION_HOST_START_ROW, TRANSLATION_HOST_START_COLUMN, True

    PrepareRegistryFixture
End Sub

Private Function BuildImportWorkbookFixture() As Workbook
    Dim wb As Workbook
    Dim dictSheet As Worksheet
    Dim exportSheet As Worksheet
    Dim analysisSheet As Worksheet
    Dim translationSheet As Worksheet

    Set wb = TestHelpers.NewWorkbook

    Set dictSheet = TestHelpers.EnsureWorksheet(DICTIONARY_SHEET_NAME, wb)
    PopulateDictionarySheet dictSheet, SOURCE_DICTIONARY_VARIABLE, "ImportSheet", SOURCE_START_ROW, SOURCE_START_COLUMN

    Set exportSheet = TestHelpers.EnsureWorksheet(EXPORTS_SHEET_NAME, wb)
    PopulateExportsSheet exportSheet, SOURCE_EXPORT_STATUS, SOURCE_EXPORT_FILE_NAME, SOURCE_EXPORT_LABEL, SOURCE_START_ROW, SOURCE_START_COLUMN

    Set analysisSheet = TestHelpers.EnsureWorksheet(ANALYSIS_SHEET_NAME, wb)
    PopulateAnalysisSheet analysisSheet, "Import", SOURCE_ANALYSIS_HEADER

    Set translationSheet = TestHelpers.EnsureWorksheet(TRANSLATIONS_SHEET_NAME, wb)
    PopulateTranslationsSheet translationSheet, "Import label", SOURCE_TRANSLATION_VALUE, SOURCE_TRANSLATION_TAG, _
                             TRANSLATION_SOURCE_START_ROW, TRANSLATION_SOURCE_START_COLUMN, False

    On Error Resume Next
        wb.Names("__ll_exports_total__").Delete
    On Error GoTo 0
    wb.Names.Add Name:="__ll_exports_total__", RefersTo:="=2"

    Set BuildImportWorkbookFixture = wb
End Function

Private Sub PopulateDictionarySheet(ByVal targetSheet As Worksheet, _
                                    ByVal variableName As String, _
                                    ByVal sheetValue As String, _
                                    ByVal startRow As Long, _
                                    ByVal startColumn As Long)

    Dim headers As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    headers = DictionaryHeaders()
    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow, startColumn), headerMatrix

    dataMatrix = TestHelpers.RowsToMatrix(Array(BuildDictionaryDataRow(variableName, sheetValue)))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow + 1, startColumn), dataMatrix
End Sub

Private Function DictionaryHeaders() As Variant
    DictionaryHeaders = Split(DICTIONARY_HEADERS_DEFINITION, "|")
End Function

Private Function BuildDictionaryDataRow(ByVal variableName As String, _
                                        ByVal sheetValue As String) As Variant

    Dim headers As Variant
    Dim values() As Variant
    Dim idx As Long
    Dim headerText As String

    headers = DictionaryHeaders()
    ReDim values(LBound(headers) To UBound(headers))

    For idx = LBound(headers) To UBound(headers)
        headerText = LCase$(CStr(headers(idx)))
        Select Case headerText
            Case "variable name"
                values(idx) = variableName
            Case "main label"
                values(idx) = variableName & " label"
            Case "sheet name"
                values(idx) = sheetValue
            Case "sheet type"
                values(idx) = "hlist2D"
            Case "status"
                values(idx) = "active"
            Case "control"
                values(idx) = "text"
            Case "unique"
                values(idx) = "no"
            Case Else
                values(idx) = vbNullString
        End Select
    Next idx

    BuildDictionaryDataRow = values
End Function

Private Sub PopulateExportsSheet(ByVal targetSheet As Worksheet, _
                                 ByVal statusValue As String, _
                                 ByVal fileNameValue As String, _
                                 ByVal labelValue As String, _
                                 ByVal startRow As Long, _
                                 ByVal startColumn As Long)

    Dim headers As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim totalColumns As Long
    Dim dataRows As Long
    Dim sourceRange As Range
    Dim lo As ListObject

    headers = ExportHeaders()
    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow, startColumn), headerMatrix

    dataMatrix = TestHelpers.RowsToMatrix(Array(BuildExportDataRow(statusValue, fileNameValue, labelValue)))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow + 1, startColumn), dataMatrix

    totalColumns = UBound(headers) - LBound(headers) + 1
    dataRows = UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 1

    Set sourceRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                        targetSheet.Cells(startRow + dataRows, startColumn + totalColumns - 1))

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
    lo.Name = "TST_Exports"
    lo.TableStyle = ""
End Sub

Private Function ExportHeaders() As Variant
    ExportHeaders = Array( _
        "export number", _
        "status", _
        "label button", _
        "file format", _
        "file name", _
        "password", _
        "include personal identifiers", _
        "include p-codes", _
        "header format", _
        "export metadata sheets", _
        "export analyses sheets")
End Function

Private Function BuildExportDataRow(ByVal statusValue As String, _
                                    ByVal fileNameValue As String, _
                                    ByVal labelValue As String) As Variant

    Dim headers As Variant
    Dim values() As Variant
    Dim idx As Long
    Dim headerText As String

    headers = ExportHeaders()
    ReDim values(LBound(headers) To UBound(headers))

    For idx = LBound(headers) To UBound(headers)
        headerText = LCase$(CStr(headers(idx)))
        Select Case headerText
            Case "export number"
                values(idx) = 1
            Case "status"
                values(idx) = statusValue
            Case "label button"
                values(idx) = labelValue
            Case "file format"
                values(idx) = "xlsx"
            Case "file name"
                values(idx) = fileNameValue
            Case "password"
                values(idx) = "pwd"
            Case "include personal identifiers"
                values(idx) = "yes"
            Case "include p-codes"
                values(idx) = "no"
            Case "header format"
                values(idx) = "default"
            Case "export metadata sheets", "export analyses sheets"
                values(idx) = "no"
            Case Else
                values(idx) = vbNullString
        End Select
    Next idx

    BuildExportDataRow = values
End Function

Private Sub PopulateTranslationsSheet(ByVal targetSheet As Worksheet, _
                                      ByVal labelValue As String, _
                                      ByVal translationValue As String, _
                                      ByVal tagValue As String, _
                                      ByVal startRow As Long, _
                                      ByVal startColumn As Long, _
                                      Optional ByVal includeTagColumn As Boolean = True)

    Dim lo As ListObject
    Dim headerRange As Range

    targetSheet.Cells.Clear

    targetSheet.Cells(startRow, startColumn).Value = "label"
    targetSheet.Cells(startRow, startColumn + 1).Value = "English"
    targetSheet.Cells(startRow + 1, startColumn).Value = labelValue
    targetSheet.Cells(startRow + 1, startColumn + 1).Value = translationValue

    Set headerRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                        targetSheet.Cells(startRow + 1, startColumn + 1))

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
    lo.Name = TRANSLATIONS_TABLE_NAME
    lo.TableStyle = ""

    If includeTagColumn Then
        targetSheet.Cells(startRow + 1, startColumn - 1).Value = tagValue
        With SetupTranslationsTable.Create(lo)
            .SetDisplayPrompts False
        End With
    End If
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
End Sub

Private Sub PopulateAnalysisSheet(ByVal targetSheet As Worksheet, _
                                  ByVal prefix As String, _
                                  ByVal headerText As String)

    Dim nextRow As Long

    targetSheet.Cells.Clear
    targetSheet.Cells(2, 1).Value = headerText
    nextRow = 3

    nextRow = AddAnalysisTable(targetSheet, nextRow, "Tab_global_summary", _
                               Array("Section"), _
                               Array(Array(prefix & " global section")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_Univariate_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " univariate section")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_Bivariate_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " bivariate section")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_TimeSeries_Analysis", _
                               Array("Table order", "Section", "series id"), _
                               Array(Array(1, prefix & " timeseries one", prefix & "_series_1"), _
                                     Array(2, prefix & " timeseries two", prefix & "_series_2")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_Spatial_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " spatial section")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_Graph_TimeSeries", _
                               Array("Graph ID", "Section"), _
                               Array(Array(prefix & "_graph_1", prefix & " graph section"), _
                                     Array(prefix & "_graph_2", prefix & " graph section"), _
                                     Array(prefix & "_graph_3", prefix & " graph section"), _
                                     Array(prefix & "_graph_4", prefix & " graph section")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_Label_TSGraph", _
                               Array("Graph ID", "Section"), _
                               Array(Array(prefix & "_graph_title", prefix & " graph title")))

    nextRow = AddAnalysisTable(targetSheet, nextRow + 2, "Tab_SpatioTemporal_Analysis", _
                               Array("Section (select)"), _
                               Array(Array(prefix & " spatio one"), _
                                     Array(prefix & " spatio two"), _
                                     Array(prefix & " spatio three")))

    Call AddAnalysisTable(targetSheet, nextRow + 2, "Tab_SpatioTemporal_Specs", _
                          Array("Section"), _
                          Array(Array(prefix & " spatio specs")))
End Sub

Private Function AddAnalysisTable(ByVal targetSheet As Worksheet, _
                                  ByVal startRow As Long, _
                                  ByVal tableName As String, _
                                  ByVal headers As Variant, _
                                  ByVal dataRows As Variant) As Long

    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim totalColumns As Long
    Dim totalDataRows As Long
    Dim loRange As Range
    Dim lo As ListObject

    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow, 1), headerMatrix

    dataMatrix = TestHelpers.RowsToMatrix(dataRows)
    TestHelpers.WriteMatrix targetSheet.Cells(startRow + 1, 1), dataMatrix

    totalColumns = UBound(headers) - LBound(headers) + 1
    totalDataRows = UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 1

    Set loRange = targetSheet.Range(targetSheet.Cells(startRow, 1), _
                                    targetSheet.Cells(startRow + totalDataRows, totalColumns))

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, loRange, , xlYes)
    lo.Name = tableName
    lo.TableStyle = ""

    AddAnalysisTable = loRange.Row + loRange.Rows.Count + 2
End Function

Private Sub AssertImportedDictionary()
    Dim dictSheet As Worksheet
    Dim variableName As String
    Dim exportTotal As Long

    Set dictSheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
    variableName = CStr(dictSheet.Cells(DICTIONARY_HOST_START_ROW + 1, DICTIONARY_HOST_START_COLUMN).Value)

    Assert.AreEqual SOURCE_DICTIONARY_VARIABLE, variableName, "Dictionary import should replace the variable name."

    exportTotal = HostExportTotal()
    Assert.AreEqual CLng(2), exportTotal, "Dictionary import should update the export counter."
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

    Assert.AreEqual SOURCE_ANALYSIS_HEADER, CStr(analysisSheet.Cells(2, 1).Value), _
                    "Analysis import should copy the header helper cell."

    Set summaryTable = analysisSheet.ListObjects("Tab_global_summary")
    Assert.AreEqual "Import global section", _
                    CStr(summaryTable.DataBodyRange.Cells(1, 1).Value), _
                    "Analysis import should copy table rows."
End Sub

Private Sub AssertImportedTranslations()
    Dim translationSheet As Worksheet
    Dim lo As ListObject
    Dim labelIdx As Long
    Dim firstTag As String

    Set translationSheet = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
    Set lo = translationSheet.ListObjects(TRANSLATIONS_TABLE_NAME)

    labelIdx = lo.ListColumns("label").Index
    Assert.AreEqual SOURCE_TRANSLATION_VALUE, _
                    CStr(lo.DataBodyRange.Cells(1, labelIdx).Value), _
                    "Translations import should load labels from the registry watchers."

    'Ensure headers from the source workbook are preserved.
    Assert.AreEqual "English", lo.ListColumns("English").Name, _
                    "Translations import should keep existing headers."

    Assert.AreEqual CLng(2), CLng(lo.ListRows.Count), _
                    "Translations import should rebuild the table based on registry ranges."

    firstTag = CStr(lo.DataBodyRange.Cells(1, 1).Offset(0, -1).Value)
    Assert.IsTrue InStr(1, firstTag, REGISTRY_RANGE_NAME, vbTextCompare) > 0, _
                   "Translations import should assign registry-based tags."

    Assert.AreEqual CLng(1), RegistryCounterValue(), _
                    "Translations registry refresh should increment the counter."
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
