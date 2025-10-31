Attribute VB_Name = "SetupImportTestFixture"
Attribute VB_Description = "Helpers creating real setup worksheets for SetupImportService tests"

Option Explicit

'@Folder("Tests.Helpers")
'@ModuleDescription("Helpers creating real setup worksheets for SetupImportService tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@section Dictionary helpers
'===============================================================================
'@sub-title Populate a dictionary worksheet at the specified location.
'@param sheetName String name of the worksheet to populate.
'@param variableName String variable identifier stored in the first data row.
'@param sheetValue String sheet label recorded in the dictionary row.
'@param startRow Long starting row for the dictionary headers.
'@param startColumn Long starting column for the dictionary headers.
'@param targetBook Optional Workbook hosting the worksheet.
Public Sub PrepareSetupDictionarySheet(ByVal sheetName As String, _
                                       ByVal variableName As String, _
                                       ByVal sheetValue As String, _
                                       ByVal startRow As Long, _
                                       ByVal startColumn As Long, _
                                       Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim headers As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb, clearSheet:=True)

    headers = DictionaryTestFixture.DictionaryFixtureHeaders()
    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    dataMatrix = TestHelpers.RowsToMatrix(Array(BuildDictionaryRow(headers, variableName, sheetValue)))

    TestHelpers.WriteMatrix sh.Cells(startRow, startColumn), headerMatrix
    TestHelpers.WriteMatrix sh.Cells(startRow + 1, startColumn), dataMatrix
End Sub

Private Function BuildDictionaryRow(ByVal headers As Variant, _
                                    ByVal variableName As String, _
                                    ByVal sheetValue As String) As Variant
    Dim values() As Variant
    Dim idx As Long
    Dim headerText As String

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

    BuildDictionaryRow = values
End Function


'@section Choices helpers
'===============================================================================
'@sub-title Populate a choices worksheet and anchor the headers at the desired location.
'@param sheetName String name of the choices worksheet to reset.
'@param startRow Long row index where headers should begin.
'@param startColumn Long column index where headers should begin.
'@param targetBook Optional Workbook hosting the worksheet.
Public Sub PrepareSetupChoicesSheet(ByVal sheetName As String, _
                                    ByVal startRow As Long, _
                                    ByVal startColumn As Long, _
                                    Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim headers As Variant
    Dim dataRows As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb, clearSheet:=True)

    headers = ChoicesTestFixture.ChoicesFixtureHeaders()
    dataRows = ChoicesSampleRows()

    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    dataMatrix = TestHelpers.RowsToMatrix(dataRows)

    TestHelpers.WriteMatrix sh.Cells(startRow, startColumn), headerMatrix
    TestHelpers.WriteMatrix sh.Cells(startRow + 1, startColumn), dataMatrix
End Sub

Private Function ChoicesSampleRows() As Variant
    ChoicesSampleRows = Array( _
        Array("list_primary", 1, "Choice A", "Short A"), _
        Array("list_primary", 2, "Choice B", "Short B"), _
        Array("list_secondary", 1, "Option 1", "Opt1"), _
        Array("list_secondary", 2, "Option 2", "Opt2"))
End Function


'@section Exports helpers
'===============================================================================
'@sub-title Populate an exports worksheet and build the associated table.
'@param sheetName String name of the exports worksheet to reset.
'@param statusValue String status written to the export row.
'@param fileNameValue String file name captured in the export row.
'@param labelValue String label written to the export row.
'@param startRow Long starting row for the exports table.
'@param startColumn Long starting column for the exports table.
'@param targetBook Optional Workbook hosting the worksheet.
Public Sub PrepareSetupExportsSheet(ByVal sheetName As String, _
                                    ByVal statusValue As String, _
                                    ByVal fileNameValue As String, _
                                    ByVal labelValue As String, _
                                    ByVal startRow As Long, _
                                    ByVal startColumn As Long, _
                                    Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim headers As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim totalColumns As Long
    Dim dataRows As Long
    Dim sourceRange As Range
    Dim lo As ListObject

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb, clearSheet:=True)

    headers = ExportHeaders()
    headerMatrix = TestHelpers.RowsToMatrix(Array(headers))
    dataMatrix = TestHelpers.RowsToMatrix(Array(BuildExportRow(statusValue, fileNameValue, labelValue)))

    TestHelpers.WriteMatrix sh.Cells(startRow, startColumn), headerMatrix
    TestHelpers.WriteMatrix sh.Cells(startRow + 1, startColumn), dataMatrix

    totalColumns = UBound(headers) - LBound(headers) + 1
    dataRows = UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 1

    Set sourceRange = sh.Range(sh.Cells(startRow, startColumn), _
                               sh.Cells(startRow + dataRows, startColumn + totalColumns - 1))

    Set lo = sh.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
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

Private Function BuildExportRow(ByVal statusValue As String, _
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

    BuildExportRow = values
End Function


'@section Analysis helpers
'===============================================================================
'@sub-title Prepare an analysis worksheet seeded with fixture data.
'@param sheetName String name of the analysis worksheet to reset.
'@param prefix String prefix applied to generated rows.
'@param headerText String header displayed above the tables.
'@param targetBook Optional Workbook hosting the worksheet.
Public Sub PrepareSetupAnalysisSheet(ByVal sheetName As String, _
                                     ByVal prefix As String, _
                                     ByVal headerText As String, _
                                     Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim nextRow As Long

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb, clearSheet:=True)

    sh.Cells(2, 1).Value = headerText
    nextRow = 3

    nextRow = AddAnalysisTable(sh, nextRow, "Tab_global_summary", _
                               Array("Section"), _
                               Array(Array(prefix & " global section")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_Univariate_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " univariate section")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_Bivariate_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " bivariate section")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_TimeSeries_Analysis", _
                               Array("Table order", "Section", "series id"), _
                               Array(Array(1, prefix & " timeseries one", prefix & "_series_1"), _
                                     Array(2, prefix & " timeseries two", prefix & "_series_2")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_Spatial_Analysis", _
                               Array("Section"), _
                               Array(Array(prefix & " spatial section")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_Graph_TimeSeries", _
                               Array("Graph ID", "Section"), _
                               Array(Array(prefix & "_graph_1", prefix & " graph section"), _
                                     Array(prefix & "_graph_2", prefix & " graph section"), _
                                     Array(prefix & "_graph_3", prefix & " graph section"), _
                                     Array(prefix & "_graph_4", prefix & " graph section")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_Label_TSGraph", _
                               Array("Graph ID", "Section"), _
                               Array(Array(prefix & "_graph_title", prefix & " graph title")))

    nextRow = AddAnalysisTable(sh, nextRow + 2, "Tab_SpatioTemporal_Analysis", _
                               Array("Section (select)"), _
                               Array(Array(prefix & " spatio one"), _
                                     Array(prefix & " spatio two"), _
                                     Array(prefix & " spatio three")))

    Call AddAnalysisTable(sh, nextRow + 2, "Tab_SpatioTemporal_Specs", _
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

    Set lo = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                         Source:=loRange, _
                                         XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
    lo.TableStyle = ""

    AddAnalysisTable = startRow + totalDataRows + 1
End Function


'@section Translations helpers
'===============================================================================
'@sub-title Seed the translations worksheet and corresponding listobject.
'@param sheetName String name of the translations worksheet.
'@param tableName String name assigned to the ListObject.
'@param labelValue String label stored in the first data row.
'@param translationValue String translation stored alongside the label.
'@param tagValue String tag persisted when includeTagColumn is True.
'@param startRow Long starting row for the translation headers.
'@param startColumn Long starting column for the translation headers.
'@param includeTagColumn Optional Boolean enabling the tag column.
'@param targetBook Optional Workbook hosting the worksheet.
Public Sub PrepareSetupTranslationsSheet(ByVal sheetName As String, _
                                         ByVal tableName As String, _
                                         ByVal labelValue As String, _
                                         ByVal translationValue As String, _
                                         ByVal tagValue As String, _
                                         ByVal startRow As Long, _
                                         ByVal startColumn As Long, _
                                         Optional ByVal includeTagColumn As Boolean = True, _
                                         Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim lo As ListObject
    Dim matrixRows As Long
    Dim matrixCols As Long
    Dim sourceRange As Range

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb, clearSheet:=True)

    headerMatrix = TestHelpers.RowsToMatrix(Array(Array("Lang1", "English")))
    dataMatrix = TestHelpers.RowsToMatrix(Array(Array(labelValue, translationValue)))

    TestHelpers.WriteMatrix sh.Cells(startRow, startColumn), headerMatrix
    TestHelpers.WriteMatrix sh.Cells(startRow + 1, startColumn), dataMatrix

    matrixRows = (UBound(headerMatrix, 1) - LBound(headerMatrix, 1) + 1) + _
                 (UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 1)
    matrixCols = UBound(headerMatrix, 2) - LBound(headerMatrix, 2) + 1

    Set sourceRange = sh.Range(sh.Cells(startRow, startColumn), _
                               sh.Cells(startRow + matrixRows - 1, startColumn + matrixCols - 1))

    Set lo = sh.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
    lo.Name = tableName
    lo.TableStyle = ""

    If includeTagColumn Then
        sh.Cells(startRow + 1, startColumn - 1).Value = tagValue
        With SetupTranslationsTable.Create(lo)
            .SetDisplayPrompts False
        End With
    End If
End Sub


'@section Workbook resolver
'===============================================================================
Private Function ResolveWorkbook(ByVal candidate As Workbook) As Workbook
    If candidate Is Nothing Then
        Set ResolveWorkbook = ThisWorkbook
    Else
        Set ResolveWorkbook = candidate
    End If
End Function
