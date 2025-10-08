Attribute VB_Name = "LLFormatTestFixture"
Attribute VB_Description = "Helpers for LLFormat tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Helpers for LLFormat tests")
'@details Seeds LL format worksheets used by tests with hard-coded values and colour formatting.

Public Const LLFORMAT_TEMPLATE_SHEET As String = "LLFormatFixture"

Private Const TABLE_NAME As String = "LLFormatTable"
Private formatHeaders As Variant
Private formatRows As Variant

Private Const ROW_SCOPE_INDEX As Long = 0
Private Const ROW_LABEL_INDEX As Long = 1
Private Const ROW_DESIGN1_VALUE_INDEX As Long = 2
Private Const ROW_DESIGN1_FONT_INDEX As Long = 3
Private Const ROW_DESIGN1_INTERIOR_INDEX As Long = 4
Private Const ROW_DESIGN2_VALUE_INDEX As Long = 5
Private Const ROW_DESIGN2_INTERIOR_INDEX As Long = 6
Private Const ROW_DESIGN2_FONT_INDEX As Long = 7


'@section Fixture Data
'===============================================================================

Private Sub EnsureFormatFixtureData()
    If IsEmpty(formatHeaders) Then formatHeaders = FormatHeadersArray()
    If IsEmpty(formatRows) Then formatRows = FormatRowsArray()
End Sub

Private Function FormatHeadersArray() As Variant
    FormatHeadersArray = Array("scope", "label", "design 1", "design 2")
End Function

Private Function FormatRowsArray() As Variant
    Dim rows As Collection

    Set rows = New Collection

    AddFormatRow rows, "Linelist Analysis", "Analysis base font size", 11, 1331390, 16777215, 11, 16777215, 1331390
    AddFormatRow rows, "Linelist Analysis, all", "Analysis Table borders color", Empty, 0, 16719904, Empty, 4275238, 1331390
    AddFormatRow rows, "Linelist Hlist, Vlist And analysis", "Button default font color", Empty, 0, 16777215, Empty, 16777215, 0
    AddFormatRow rows, "Linelist Hlist, Vlist And analysis", "Button default interior color", Empty, 0, 11892015, Empty, 78168, 1331390
    AddFormatRow rows, "Linelist Hlist", "Calculated Formula column interior color", Empty, 1331390, 16777215, Empty, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist", "Calculated formula font color", Empty, 0, 9147815, Empty, 9147815, 0
    AddFormatRow rows, "Linelist Hlist", "Calculated formula header color", Empty, 0, 15132391, Empty, 15132391, 0
    AddFormatRow rows, "Linelist Analysis, all", "Categories names font color", Empty, 0, 9109504, Empty, 78168, 1331390
    AddFormatRow rows, "Linelist Analysis, all", "default analysis column width", 20, 1331390, 16777215, 20, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist, Linelist Vlist", "default linelist column width", 22, 1331390, 16777215, 22, 16777215, 1331390
    AddFormatRow rows, "Linelist Analysis, all", "default row height gap for graphs", 45, 1331390, 16777215, 45, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist", "Entry Table Style", "None, with borders", 0, 16777215, "None, with borders", 16777215, 0
    AddFormatRow rows, "Linelist Hlist", "Geo font color", Empty, 0, 5855577, Empty, 6375440, 6375440
    AddFormatRow rows, "Linelist Hlist", "Geo interior color", Empty, 0, 11389944, Empty, 11389944, 0
    AddFormatRow rows, "Linelist Vlist, Linelist Hlist", "hlist and vlist table borders color", Empty, 0, 16719904, Empty, 4275238, 1331390
    AddFormatRow rows, "Linelist Hlist", "Hlist main label font color", Empty, 0, 0, Empty, 4275238, 1331390
    AddFormatRow rows, "Linelist Hlist", "Hlist main label font size", 9, 1331390, 16777215, 9, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist", "Hlist sub label font color", Empty, 1331390, 11892015, Empty, 78168, 1331390
    AddFormatRow rows, "Linelist Hlist", "Hlist sub label font size", 8, 1331390, 16777215, 8, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist", "Hlist Table Borders color", Empty, 0, 16719904, Empty, 4275238, 1331390
    AddFormatRow rows, "Linelist Hlist", "hlist table header color", Empty, 15789801, 15789801, Empty, 2102022, 0
    AddFormatRow rows, "Linelist Hlist, Vlist", "Linelist base font size", 9, 1331390, 16777215, 9, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist, Vlist", "main section font color", Empty, 0, 0, Empty, 16777215, 0
    AddFormatRow rows, "Linelist Hlist, Linelist Vlist", "main section font size", 12, 1331390, 16777215, 12, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist, Vlist", "main section interior color", Empty, 0, 14723184, Empty, 78168, 0
    AddFormatRow rows, "Linelist Analysis, all", "missing font color", Empty, 0, 7952452, Empty, 2102022, 0
    AddFormatRow rows, "Linelist Analysis, all", "missing interior Color", Empty, 0, 15789801, Empty, 13684944, 0
    AddFormatRow rows, "Linelist Hlist", "Notes font color", Empty, 0, 5729755, Empty, 2102022, 0
    AddFormatRow rows, "Linelist Hlist, Vlist And analysis", "Select dropdown font color", Empty, 0, 16777215, Empty, 16777215, 16777215
    AddFormatRow rows, "Linelist Hlist, Vlist And analysis", "Select dropdown interior color", Empty, 0, 11892015, Empty, 78168, 0
    AddFormatRow rows, "Linelist Hlist, Linelist Vlist", "sub section font color", Empty, 0, 11892015, Empty, 78168, 0
    AddFormatRow rows, "Linelist Hlist, Linelist Vlist", "sub section font size", 11, 1331390, 16777215, 11, 16777215, 1331390
    AddFormatRow rows, "Linelist Hlist, Linelist Vlist", "sub section interior color", Empty, 0, 16247773, Empty, 16777215, 0
    AddFormatRow rows, "Linelist Analysis, all", "Table categories font color", Empty, 0, 9109504, Empty, 4145464, 0
    AddFormatRow rows, "Linelist Analysis, all", "Table categories interior color", Empty, 0, 16775664, Empty, 15921906, 0
    AddFormatRow rows, "Linelist Analysis, all", "Table sections font color", Empty, 0, 9109504, Empty, 78168, 0
    AddFormatRow rows, "Linelist Analysis, all", "Table sections interior color", Empty, 0, 16777215, Empty, 16777215, 0
    AddFormatRow rows, "Linelist Analysis, all", "Table title font color", Empty, 0, 16719904, Empty, 78168, 0
    AddFormatRow rows, "Linelist Analysis, time series", "Time series header font color", Empty, 0, 16777215, Empty, 16777215, 0
    AddFormatRow rows, "Linelist Analysis, time series", "Time series header interior color", Empty, 0, 11892015, Empty, 78168, 0
    AddFormatRow rows, "Linelist Vlist", "Vlist main label font color", Empty, 0, 9917741, Empty, 3564453, 0
    AddFormatRow rows, "Linelist VList", "Vlist main label font size", 12, 1331390, 16777215, 12, 16777215, 1331390
    AddFormatRow rows, "Linelist VList", "Vlist sub label font color", Empty, 1331390, 8421504, Empty, 8421504, 1331390
    AddFormatRow rows, "Linelist VList", "Vlist sub label font size", 8, 1331390, 16777215, 8, 16777215, 1331390

    FormatRowsArray = RowsCollectionToVariant(rows)
End Function

Private Sub AddFormatRow(ByVal rows As Collection, _
                         ByVal scopeText As String, _
                         ByVal labelText As String, _
                         ByVal design1Value As Variant, _
                         ByVal design1FontColour As Variant, _
                         ByVal design1InteriorColour As Variant, _
                         ByVal design2Value As Variant, _
                         ByVal design2FontColour As Variant, _
                         ByVal design2InteriorColour As Variant)
    rows.Add CreateFormatRow(scopeText, labelText, design1Value, design1FontColour, design1InteriorColour, _
                              design2Value, design2FontColour, design2InteriorColour)
End Sub

Private Function RowsCollectionToVariant(ByVal rows As Collection) As Variant
    Dim result() As Variant
    Dim index As Long

    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then
        RowsCollectionToVariant = Array()
        Exit Function
    End If

    ReDim result(0 To rows.Count - 1)
    For index = 1 To rows.Count
        result(index - 1) = rows(index)
    Next index

    RowsCollectionToVariant = result
End Function

Private Function CreateFormatRow(ByVal scopeText As String, _
                                 ByVal labelText As String, _
                                 ByVal design1Value As Variant, _
                                 ByVal design1FontColour As Variant, _
                                 ByVal design1InteriorColour As Variant, _
                                 ByVal design2Value As Variant, _
                                 ByVal design2FontColour As Variant, _
                                 ByVal design2InteriorColour As Variant) As Variant

    Dim rowData(ROW_SCOPE_INDEX To ROW_DESIGN2_FONT_INDEX) As Variant

    rowData(ROW_SCOPE_INDEX) = scopeText
    rowData(ROW_LABEL_INDEX) = labelText
    rowData(ROW_DESIGN1_VALUE_INDEX) = design1Value
    rowData(ROW_DESIGN1_FONT_INDEX) = design1FontColour
    rowData(ROW_DESIGN1_INTERIOR_INDEX) = design1InteriorColour
    rowData(ROW_DESIGN2_VALUE_INDEX) = design2Value
    rowData(ROW_DESIGN2_FONT_INDEX) = design2FontColour
    rowData(ROW_DESIGN2_INTERIOR_INDEX) = design2InteriorColour

    CreateFormatRow = rowData
End Function

'@section Public API
'===============================================================================

'@description Create or rebuild the LL format fixture worksheet.
'@param sheetName String name for the worksheet to populate.
'@param targetBook Optional Workbook host; defaults to ThisWorkbook.
'@return Worksheet containing the LL format table.
Public Function PrepareLLFormatFixture(ByVal sheetName As String, _
                                       Optional ByVal targetBook As Workbook) As Worksheet

    Dim wb As Workbook
    Dim fixtureSheet As Worksheet

    On Error GoTo Fail

    Set wb = ResolveWorkbook(targetBook)

    DeleteLLFormatFixture sheetName, wb
    Set fixtureSheet = TestHelpers.EnsureWorksheet(sheetName, wb)

    BuildFormatTemplate fixtureSheet

    Set PrepareLLFormatFixture = fixtureSheet
    Exit Function

Fail:
    Err.Raise Err.Number, "LLFormatTestFixture.PrepareLLFormatFixture", Err.Description
End Function

'@description Delete an existing LL format fixture worksheet when present.
'@param sheetName String name of the worksheet to remove.
'@param targetBook Optional Workbook host; defaults to ThisWorkbook.
Public Sub DeleteLLFormatFixture(ByVal sheetName As String, _
                                 Optional ByVal targetBook As Workbook)

    Dim wb As Workbook

    Set wb = ResolveWorkbook(targetBook)

    If Not TestHelpers.WorksheetExists(sheetName, wb) Then Exit Sub

    If wb Is ThisWorkbook Then
        TestHelpers.DeleteWorksheet sheetName
    Else
        DeleteWorksheetInternal sheetName, wb
    End If
End Sub

'@description Ensure the reusable LL format template exists and is prepared.
'@param targetBook Optional Workbook host; defaults to ThisWorkbook.
'@return Worksheet reference to the LLFormat fixture template.
Public Function LLFormatTemplate(Optional ByVal targetBook As Workbook) As Worksheet

    Dim wb As Workbook
    Dim template As Worksheet

    Set wb = ResolveWorkbook(targetBook)
    Set template = TestHelpers.EnsureWorksheet(LLFORMAT_TEMPLATE_SHEET, wb)
    BuildFormatTemplate template
    Set LLFormatTemplate = template
End Function

'@description Retrieve the cell within the fixture table matching a label/design.
'@param hostSheet Worksheet containing the fixture table.
'@param labelText Label to locate.
'@param designColumn Optional design column name; defaults to the first design.
'@return Range pointing to the requested fixture cell.
Public Function FixtureCell(ByVal hostSheet As Worksheet, _
                            ByVal labelText As String, _
                            ByVal designColumn As String) As Range

    Dim tableObj As ListObject
    Dim labelRange As Range
    Dim labelCell As Range
    Dim designRange As Range
    Dim columnIndex As Long

    If hostSheet Is Nothing Then
        Err.Raise vbObjectError + 514, "LLFormatTestFixture.FixtureCell", _
                  "Fixture sheet reference is required before locating a cell"
    End If

    Set tableObj = FixtureTable(hostSheet)
    Set labelRange = tableObj.ListColumns("label").DataBodyRange

    On Error Resume Next
        Set labelCell = labelRange.Find(What:=labelText, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If labelCell Is Nothing Then
        Err.Raise vbObjectError + 516, "LLFormatTestFixture.FixtureCell", _
                  "Label '" & labelText & "' is missing from fixture sheet '" & hostSheet.Name & "'"
    End If

    designColumn = ResolveDesignName(tableObj, designColumn)
    Set designRange = tableObj.ListColumns(designColumn).Range

    columnIndex = designRange.Column
    Set FixtureCell = hostSheet.Cells(labelCell.Row, columnIndex)
End Function

'@description Return the collection of available design column names.
'@param hostSheet Worksheet containing the fixture table.
'@return Collection of design name strings.
Public Function DesignNames(ByVal hostSheet As Worksheet) As Collection
    Dim tableObj As ListObject
    Dim col As ListColumn
    Dim names As Collection

    Set tableObj = FixtureTable(hostSheet)
    Set names = New Collection

    For Each col In tableObj.ListColumns
        If StrComp(col.Name, "label", vbTextCompare) <> 0 Then
            names.Add col.Name
        End If
    Next col

    Set DesignNames = names
End Function

'@description Retrieve the default design column name for a sheet.
'@param hostSheet Worksheet containing the fixture table.
'@return String name of the default design column.
Public Function DefaultDesignName(ByVal hostSheet As Worksheet) As String
    DefaultDesignName = ResolveDesignName(FixtureTable(hostSheet), vbNullString)
End Function

'@description Get the configured interior colour for a label/design combination.
'@param hostSheet Worksheet containing the fixture table.
'@param labelText Label whose colour is required.
'@param designColumn Optional design column; defaults to the first design.
'@return Long colour value taken from the cell interior.
Public Function DesignColour(ByVal hostSheet As Worksheet, _
                             ByVal labelText As String, _
                             Optional ByVal designColumn As String = vbNullString) As Long

    Dim resolvedDesign As String
    resolvedDesign = ResolveDesignName(FixtureTable(hostSheet), designColumn)
    DesignColour = CLng(FixtureCell(hostSheet, labelText, resolvedDesign).Interior.Color)
End Function

'@description Get the numeric or textual value stored for a label/design combination.
'@param hostSheet Worksheet containing the fixture table.
'@param labelText Label whose value is required.
'@param designColumn Optional design column; defaults to the first design.
'@return Variant containing the cell value.
Public Function DesignNumericValue(ByVal hostSheet As Worksheet, _
                                   ByVal labelText As String, _
                                   Optional ByVal designColumn As String = vbNullString) As Variant

    Dim resolvedDesign As String
    resolvedDesign = ResolveDesignName(FixtureTable(hostSheet), designColumn)
    DesignNumericValue = FixtureCell(hostSheet, labelText, resolvedDesign).Value
End Function

'@section Internal helpers
'===============================================================================

Private Function ResolveWorkbook(Optional ByVal targetBook As Workbook) As Workbook
    If targetBook Is Nothing Then
        Set ResolveWorkbook = ThisWorkbook
    Else
        Set ResolveWorkbook = targetBook
    End If
End Function

Private Sub DeleteWorksheetInternal(ByVal sheetName As String, ByVal wb As Workbook)

    Dim previousAlerts As Boolean
    Dim previousUpdating As Boolean

    previousAlerts = Application.DisplayAlerts
    previousUpdating = Application.ScreenUpdating

    On Error Resume Next
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        wb.Worksheets(sheetName).Delete
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousUpdating
End Sub

Private Sub EnsureDesignTypeName(ByVal targetSheet As Worksheet)
    Dim designCell As Range
    Dim localName As Name
    Dim tableObj As ListObject
    Dim targetCell As Range
    Dim defaultDesign As String

    If targetSheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set localName = targetSheet.Names("DESIGNTYPE")
    On Error GoTo 0

    If Not localName Is Nothing Then
        Set designCell = Nothing
        On Error Resume Next
            Set designCell = localName.RefersToRange
        On Error GoTo 0
        If Not designCell Is Nothing Then
            If LenB(CStr(designCell.Value)) = 0 Then
                Set tableObj = FixtureTable(targetSheet)
                defaultDesign = FirstDesignName(tableObj)
                designCell.Value = defaultDesign
            End If
            Exit Sub
        End If
    End If

    Set tableObj = FixtureTable(targetSheet)
    defaultDesign = FirstDesignName(tableObj)

    Set targetCell = targetSheet.Cells(tableObj.HeaderRowRange.Row, _
                                       tableObj.Range.Column + tableObj.ListColumns.Count + 2)

    If LenB(CStr(targetCell.Value)) = 0 Then
        targetCell.Value = defaultDesign
    End If

    On Error Resume Next
        targetSheet.Names("DESIGNTYPE").Delete
    On Error GoTo 0
    targetSheet.Names.Add Name:="DESIGNTYPE", RefersTo:=targetCell, Visible:=False
End Sub

Private Function FixtureTable(ByVal hostSheet As Worksheet) As ListObject
    If hostSheet Is Nothing Then
        Err.Raise vbObjectError + 520, "LLFormatTestFixture.FixtureTable", _
                  "Fixture sheet reference is required"
    End If
    If hostSheet.ListObjects.Count = 0 Then
        Err.Raise vbObjectError + 515, "LLFormatTestFixture.FixtureTable", _
                  "Fixture sheet '" & hostSheet.Name & "' must expose a format table"
    End If
    Set FixtureTable = hostSheet.ListObjects(1)
End Function

Private Function ResolveDesignName(ByVal tableObj As ListObject, ByVal designColumn As String) As String
    Dim columnObject As ListColumn

    If LenB(designColumn) > 0 Then
        On Error Resume Next
            Set columnObject = tableObj.ListColumns(designColumn)
        On Error GoTo 0
        If columnObject Is Nothing Then
            Err.Raise vbObjectError + 521, "LLFormatTestFixture.ResolveDesignName", _
                      "Design column '" & designColumn & "' is missing from fixture table '" & tableObj.Name & "'"
        End If
        ResolveDesignName = columnObject.Name
    Else
        ResolveDesignName = FirstDesignName(tableObj)
    End If
End Function

Private Function FirstDesignName(ByVal tableObj As ListObject) As String
    Dim col As ListColumn
    Dim labelIndex As Long

    For Each col In tableObj.ListColumns
        If StrComp(col.Name, "label", vbTextCompare) = 0 Then
            labelIndex = col.Index
            Exit For
        End If
    Next col

    If labelIndex = 0 Then
        Err.Raise vbObjectError + 522, "LLFormatTestFixture.FirstDesignName", _
                  "Column 'label' was not found in fixture table '" & tableObj.Name & "'"
    End If

    For Each col In tableObj.ListColumns
        If col.Index > labelIndex Then
            FirstDesignName = col.Name
            Exit Function
        End If
    Next col

    Err.Raise vbObjectError + 522, "LLFormatTestFixture.FirstDesignName", _
              "No design columns were found after 'label' in fixture table '" & tableObj.Name & "'"
End Function

Private Sub BuildFormatTemplate(ByVal targetSheet As Worksheet)

    Dim matrixRows() As Variant
    Dim matrix As Variant
    Dim startCell As Range
    Dim lo As ListObject
    Dim tableRange As Range
    Dim rowIndex As Long
    Dim rowData As Variant
    Dim dataBody As Range
    Dim tableRow As Range

    EnsureFormatFixtureData

    If targetSheet Is Nothing Then Exit Sub

    Do While targetSheet.ListObjects.Count > 0
        targetSheet.ListObjects(1).Delete
    Loop

    targetSheet.Cells.Clear

    ReDim matrixRows(0 To UBound(formatRows) - LBound(formatRows) + 1)
    matrixRows(0) = formatHeaders

    For rowIndex = LBound(formatRows) To UBound(formatRows)
        rowData = formatRows(rowIndex)
        matrixRows(rowIndex - LBound(formatRows) + 1) = Array( _
            CStr(rowData(ROW_SCOPE_INDEX)), _
            CStr(rowData(ROW_LABEL_INDEX)), _
            rowData(ROW_DESIGN1_VALUE_INDEX), _
            rowData(ROW_DESIGN2_VALUE_INDEX))
    Next rowIndex

    matrix = TestHelpers.RowsToMatrix(matrixRows)
    Set startCell = targetSheet.Range("A1")
    TestHelpers.WriteMatrix startCell, matrix

    Set tableRange = startCell.Resize(UBound(matrix, 1), UBound(matrix, 2))
    Set lo = targetSheet.ListObjects.Add(xlSrcRange, source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = TABLE_NAME
    On Error Resume Next
        lo.TableStyle = "TableStyleMedium2"
    On Error GoTo 0

    If lo.DataBodyRange Is Nothing Then GoTo EnsureDesignName

    Set dataBody = lo.DataBodyRange

    For rowIndex = LBound(formatRows) To UBound(formatRows)
        rowData = formatRows(rowIndex)
        Set tableRow = dataBody.Rows(rowIndex - LBound(formatRows) + 1)

        ApplyColouring tableRow.Columns(3), rowData(ROW_DESIGN1_FONT_INDEX), rowData(ROW_DESIGN1_INTERIOR_INDEX)
        ApplyColouring tableRow.Columns(4), rowData(ROW_DESIGN2_INTERIOR_INDEX), rowData(ROW_DESIGN2_FONT_INDEX)
    Next rowIndex

EnsureDesignName:
    EnsureDesignTypeName targetSheet
End Sub

Private Sub ApplyColouring(ByVal targetCell As Range, ByVal fontColorValue As Variant, ByVal interiorColorValue As Variant)
    If targetCell Is Nothing Then Exit Sub

    If Not IsEmptyVariant(interiorColorValue) Then
        targetCell.Interior.Color = CLng(interiorColorValue)
    Else
        targetCell.Interior.Pattern = xlNone
    End If

    If Not IsEmptyVariant(fontColorValue) Then
        targetCell.Font.Color = CLng(fontColorValue)
    End If
End Sub

Private Function IsEmptyVariant(ByVal candidate As Variant) As Boolean
    If VarType(candidate) = vbEmpty Then
        IsEmptyVariant = True
    ElseIf IsObject(candidate) Then
        IsEmptyVariant = candidate Is Nothing
    ElseIf VarType(candidate) = vbString Then
        IsEmptyVariant = (Len(Trim$(CStr(candidate))) = 0)
    Else
        IsEmptyVariant = False
    End If
End Function
