Attribute VB_Name = "DictionaryTestFixture"
Attribute VB_Description = "Shared dictionary fixture for tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@Folder("Tests")
'@ModuleDescription("Shared dictionary fixture for tests")

Public Const DICTIONARY_FIXTURE_LAST_COLOR As Long = 15773696 'light blue
Private Const TABLE_NAME_HEADER As String = "Table Name"
Private Const SHEET_NAME_HEADER As String = "Sheet Name"
Private Const SHEET_TYPE_HEADER As String = "Sheet Type"

'@section Fixture Cache
'===============================================================================

Private fixtureHeaders As Variant
Private fixtureRows As Variant

Private Sub EnsureFixtureLoaded()
    If IsEmpty(fixtureHeaders) Then fixtureHeaders = FixtureHeadersArray()
    If IsEmpty(fixtureRows) Then fixtureRows = FixtureRowsArray()

    If Not HeaderArrayContains(fixtureHeaders, TABLE_NAME_HEADER) Then
        fixtureHeaders = InsertTableNameHeader(fixtureHeaders)
        fixtureRows = AddTableNamesToRows(fixtureRows, fixtureHeaders)
    ElseIf NeedsTableNameValues(fixtureRows, fixtureHeaders) Then
        fixtureRows = AddTableNamesToRows(fixtureRows, fixtureHeaders)
    End If
End Sub

'@section Worksheet Preparation
'===============================================================================

'@description Seed a worksheet with the dictionary fixture data.
'@param sheetName String. Worksheet name to host the fixture.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
Public Sub PrepareDictionaryFixture(ByVal sheetName As String, Optional ByVal targetBook As Workbook)
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim lastRow As Long
    Dim lastCol As Long

    EnsureFixtureLoaded

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    Set sh = EnsureWorksheet(sheetName, wb, visibility:=xlSheetVeryhidden)

    headerMatrix = RowsToMatrix(Array(fixtureHeaders))
    WriteMatrix sh.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(fixtureRows)
    WriteMatrix sh.Cells(2, 1), dataMatrix

    lastRow = 1 + (UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 1)
    lastCol = UBound(headerMatrix, 2) - LBound(headerMatrix, 2) + 1
    sh.Cells(lastRow, lastCol).Interior.Color = DICTIONARY_FIXTURE_LAST_COLOR

End Sub

'@section Fixture Metadata
'===============================================================================

'@description Retrieve the fixture headers.
'@return Variant array of headers.
Public Function DictionaryFixtureHeaders() As Variant
    EnsureFixtureLoaded
    DictionaryFixtureHeaders = fixtureHeaders
End Function

'@description Retrieve the fixture rows.
'@return Variant array of row arrays.
Public Function DictionaryFixtureRows() As Variant
    EnsureFixtureLoaded
    DictionaryFixtureRows = fixtureRows
End Function

'@description Get the number of data rows in the fixture.
'@return Long count of rows.
Public Function DictionaryFixtureRowCount() As Long
    EnsureFixtureLoaded
    DictionaryFixtureRowCount = UBound(fixtureRows) - LBound(fixtureRows) + 1
End Function

'@description Get the number of columns in the fixture.
'@return Long count of columns.
Public Function DictionaryFixtureColumnCount() As Long
    EnsureFixtureLoaded
    DictionaryFixtureColumnCount = UBound(fixtureHeaders) - LBound(fixtureHeaders) + 1
End Function

'@section Lookup Helpers
'===============================================================================

'@description Find the header index for a column name.
'@param columnName String. Header to locate.
'@return Long index in the headers array.
Public Function DictionaryHeaderIndex(ByVal columnName As String) As Long
    Dim idx As Long

    EnsureFixtureLoaded

    For idx = LBound(fixtureHeaders) To UBound(fixtureHeaders)
        If StrComp(CStr(fixtureHeaders(idx)), columnName, vbTextCompare) = 0 Then
            DictionaryHeaderIndex = idx
            Exit Function
        End If
    Next idx

    Err.Raise vbObjectError + 2000, "DictionaryTestFixture", "Header not found: " & columnName
End Function

'@description Retrieve a single value from the fixture.
'@param rowIndex Long. Zero-based row index.
'@param columnName String. Column header to fetch.
'@return String value stored at row/column.
Public Function DictionaryFixtureValue(ByVal rowIndex As Long, ByVal columnName As String) As String
    EnsureFixtureLoaded
    DictionaryFixtureValue = CStr(fixtureRows(rowIndex)(DictionaryHeaderIndex(columnName)))
End Function

'@description Get distinct values for a column.
'@param columnName String. Header to retrieve.
'@return BetterArray of unique values.
Public Function DictionaryDistinctValues(ByVal columnName As String) As BetterArray
    Dim results As BetterArray
    Dim value As String
    Dim rowData As Variant
    Dim colIndex As Long

    EnsureFixtureLoaded

    Set results = BetterArrayFromList()
    colIndex = DictionaryHeaderIndex(columnName)

    For Each rowData In fixtureRows
        value = CStr(rowData(colIndex))
        If Not results.Includes(value) Then results.Push value
    Next rowData

    Set DictionaryDistinctValues = results
End Function

'@description Return variables whose control column matches supplied values.
'@param controls Variant array of control values to match.
'@return BetterArray of variable names.
Public Function DictionaryControlMatches(controls As Variant) As BetterArray
    Dim matches As BetterArray
    Dim rowData As Variant
    Dim controlValue As String
    Dim idx As Long
    Dim nameIndex As Long
    Dim controlIndex As Long
    Dim typeIndex As Long
    Dim sheetType As String

    EnsureFixtureLoaded

    Set matches = BetterArrayFromList()
    nameIndex = DictionaryHeaderIndex("Variable Name")
    controlIndex = DictionaryHeaderIndex("Control")
    typeIndex = DictionaryHeaderIndex("Sheet Type")

    For Each rowData In fixtureRows
        controlValue = CStr(rowData(controlIndex))
        sheetType = CStr(rowData(typeIndex))
        If StrComp(sheetType, "hlist2D", vbTextCompare) = 0 Then
            For idx = LBound(controls) To UBound(controls)
                If StrComp(controlValue, CStr(controls(idx)), vbTextCompare) = 0 Then
                    If Not matches.Includes(CStr(rowData(nameIndex))) Then matches.Push CStr(rowData(nameIndex))
                    Exit For
                End If
            Next idx
        End If
    Next rowData

    Set DictionaryControlMatches = matches
End Function

'@description Filter variables where a column equals a specific value.
'@param columnName String header to inspect.
'@param expectedValue String value to match.
'@return BetterArray of variable names.
Public Function DictionaryFieldEquals(ByVal columnName As String, ByVal expectedValue As String) As BetterArray
    Dim matches As BetterArray
    Dim rowData As Variant
    Dim nameIndex As Long
    Dim columnIndex As Long

    EnsureFixtureLoaded

    Set matches = BetterArrayFromList()
    nameIndex = DictionaryHeaderIndex("Variable Name")
    columnIndex = DictionaryHeaderIndex(columnName)

    For Each rowData In fixtureRows
        If StrComp(CStr(rowData(columnIndex)), expectedValue, vbTextCompare) = 0 Then
            If Not matches.Includes(CStr(rowData(nameIndex))) Then matches.Push CStr(rowData(nameIndex))
        End If
    Next rowData

    Set DictionaryFieldEquals = matches
End Function

'@section Fixture Data
'===============================================================================

Private Function FixtureHeadersArray() As Variant
    FixtureHeadersArray = Split("Variable Name|Main Label|Dev Comments|Editable Label|Sub Label|Note|Sheet Name|Sheet Type|Main Section|Sub Section|Status|Register Book|Personal Identifier|Variable Type|Variable Format|Control|Control Details|Unique|Export 1|Export 2|Export 3|Export 4|Export 5|Min|Max|Alert|Message|Formatting Condition|Formatting Values", "|")
End Function

Private Function FixtureRowsArray() As Variant
    FixtureRowsArray = CombineRowSets( _
        FixtureRowsChunk1(), _
        FixtureRowsChunk2(), _
        FixtureRowsChunk3(), _
        FixtureRowsChunk4())
End Function

Private Function HeaderArrayContains(ByVal headers As Variant, ByVal columnName As String) As Boolean
    Dim idx As Long
    For idx = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(idx)), columnName, vbTextCompare) = 0 Then
            HeaderArrayContains = True
            Exit Function
        End If
    Next idx
End Function

Private Function InsertTableNameHeader(ByVal headers As Variant) As Variant
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim insertAt As Long
    Dim idx As Long
    Dim result() As Variant

    lowerBound = LBound(headers)
    upperBound = UBound(headers)
    insertAt = HeaderIndexOf(headers, SHEET_TYPE_HEADER) + 1

    ReDim result(lowerBound To upperBound + 1)

    For idx = lowerBound To upperBound + 1
        If idx = insertAt Then
            result(idx) = TABLE_NAME_HEADER
        ElseIf idx < insertAt Then
            result(idx) = headers(idx)
        Else
            result(idx) = headers(idx - 1)
        End If
    Next idx

    InsertTableNameHeader = result
End Function

Private Function AddTableNamesToRows(ByVal rows As Variant, ByVal headers As Variant) As Variant
    Dim result() As Variant
    Dim idx As Long
    Dim sheetAssignments As Object
    Dim sheetIndex As Long
    Dim tableIndex As Long

    sheetIndex = HeaderIndexOf(headers, SHEET_NAME_HEADER)
    tableIndex = HeaderIndexOf(headers, TABLE_NAME_HEADER)

    ReDim result(LBound(rows) To UBound(rows))

    Set sheetAssignments = CreateObject("Scripting.Dictionary")
    sheetAssignments.CompareMode = vbTextCompare

    For idx = LBound(rows) To UBound(rows)
        result(idx) = InsertTableNameValue(rows(idx), sheetIndex, tableIndex, sheetAssignments)
    Next idx

    AddTableNamesToRows = result
End Function

Private Function InsertTableNameValue( ByVal rowValues As Variant, _
                                       ByVal sheetIndex As Long, _
                                       ByVal tableIndex As Long, _
                                       ByVal sheetAssignments As Object) As Variant
    Dim newRow() As Variant
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim idx As Long
    Dim tableName As String
    Dim sheetName As String

    lowerBound = LBound(rowValues)
    upperBound = UBound(rowValues)
    ReDim newRow(lowerBound To upperBound + 1)

    sheetName = CStr(rowValues(sheetIndex))
    tableName = ResolveTableName(sheetName, sheetAssignments)

    For idx = lowerBound To upperBound + 1
        If idx = tableIndex Then
            newRow(idx) = tableName
        ElseIf idx < tableIndex Then
            newRow(idx) = rowValues(idx)
        Else
            newRow(idx) = rowValues(idx - 1)
        End If
    Next idx

    InsertTableNameValue = newRow
End Function

Private Function ResolveTableName(ByVal sheetName As String, ByVal assignments As Object) As String
    If Not assignments.Exists(sheetName) Then assignments.Add sheetName, "table" & CStr(assignments.Count + 1)
    ResolveTableName = CStr(assignments(sheetName))
End Function

Private Function NeedsTableNameValues(ByVal rows As Variant, ByVal headers As Variant) As Boolean
    Dim anyRow As Variant
    If Not IsArray(rows) Then Exit Function
    If UBound(rows) < LBound(rows) Then Exit Function
    anyRow = rows(LBound(rows))
    NeedsTableNameValues = ArrayLength(anyRow) <> ArrayLength(headers)
End Function

Private Function HeaderIndexOf(ByVal headers As Variant, ByVal columnName As String) As Long
    Dim idx As Long
    For idx = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(idx)), columnName, vbTextCompare) = 0 Then
            HeaderIndexOf = idx
            Exit Function
        End If
    Next idx
    Err.Raise vbObjectError + 2001, "DictionaryTestFixture", "Header not found during fixture augmentation: " & columnName
End Function

Private Function ArrayLength(ByVal values As Variant) As Long
    ArrayLength = UBound(values) - LBound(values) + 1
End Function

Private Function CombineRowSets(ParamArray chunks() As Variant) As Variant
    Dim total As Long
    Dim chunk As Variant
    Dim item As Variant
    Dim result() As Variant
    Dim index As Long

    For Each chunk In chunks
        total = total + (UBound(chunk) - LBound(chunk) + 1)
    Next chunk

    ReDim result(0 To total - 1)

    For Each chunk In chunks
        For Each item In chunk
            result(index) = item
            index = index + 1
        Next item
    Next chunk

    CombineRowSets = result
End Function

Private Function FixtureRowsChunk1() As Variant
    FixtureRowsChunk1 = Array( _
        Array("hid_beg_v1", "Hidden variable at the begining", "", "", "Should be hidden", "", "vlist1D-sheet1", "vlist1D", "Hidden Section", "", "hidden", "", "", "", "", "", "", "", "6", "", "1", "12", "yes", "", "", "", "", "", ""), _
        Array("choi_v1", "Choices on vlist1D", "", "", "Values: A, B, C", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "choice_manual", "list_correct_order", "", "7", "", "2", "22", "yes", "", "", "", "", "", ""), _
        Array("choi_mult_v1", "Choice multiple on vlist1D", "", "", "Multiple values: A, B, C", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "choice_multiple", "list_multiple", "", "8", "", "3", "2", "yes", "", "", "", "", "", ""), _
        Array("choi_ord_v1", "Choice order on vlist1D", "", "", "Values: B, C, A", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "choice_manual", "list_uncorrect_order", "", "9", "", "4", "66", "yes", "", "", "", "", "", ""), _
        Array("choi_cust_v1", "Custom choices on vlist1D", "", "", "Random input by the user", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "choice_custom", "", "", "10", "", "5", "87", "yes", "", "", "", "", "", ""), _
        Array("form_v1", "Formula on vlist1D", "", "", "Formula", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "formula", "IF(ISBLANK(choi_v1), """", choi_v1 & ""-OK"")", "", "11", "", "6", "10", "yes", "", "", "", "", "", ""), _
        Array("brok_form_v1", "Broken formula on vlist1D", "", "", "This formula should fail", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "formula", "IF(ISBLANK(choi_v2), """", choi_v2 & ""-OK"")", "", "12", "", "7", "59", "", "", "", "", "", "", ""), _
        Array("hid_v1", "Hidden variable in the middle", "", "", "Should be hidden", "", "vlist1D-sheet1", "vlist1D", "Status", "", "hidden", "", "", "", "", "", "", "", "13", "", "8", "77", "yes", "", "", "", "", "", ""), _
        Array("opt_hid_v1", "Optional hidden variable on vlist1D", "", "", "Should be hidden, the user can unhide", "", "vlist1D-sheet1", "vlist1D", "Status", "", "optional, hidden", "", "", "", "", "", "", "", "14", "", "9", "72", "yes", "", "", "", "", "", ""), _
        Array("opt_vis_v1", "Optional visible variable on vlist1D", "", "", "Should be visible, the user can hide", "", "vlist1D-sheet1", "vlist1D", "Status", "", "optional, visible", "", "", "", "", "", "", "", "15", "", "10", "82", "yes", "", "", "", "", "", ""), _
        Array("mand_v1", "Mandatory variable on vlist1D", "", "", "Mandatory", "", "vlist1D-sheet1", "vlist1D", "Status", "", "mandatory", "", "", "", "", "", "", "", "16", "", "11", "60", "yes", "", "", "", "", "", ""), _
        Array("c_wh_v1", "Case when on vlist1D", "", "", "Case When formula", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "case_when", "CASE_WHEN(choi_ord_v1 = ""A"", ""Choice order is A"", choi_ord_v1 = ""B"", ""Choice order is B"", ""Unknown choice order"")", "", "17", "", "12", "56", "yes", "", "", "", "", "", ""), _
        Array("choi_form_v1", "Choice formula on vlist1D", "", "", "Choice Formula", "", "vlist1D-sheet1", "vlist1D", "Controls", "", "", "", "", "", "", "choice_formula", "CHOICE_FORMULA(list_multiple, choi_v1 = ""A"", ""choice 1"", choi_v1 = ""B"", ""choice 2"", choi_v1 = ""C"", c_wh_v1)", "", "18", "", "13", "80", "yes", "", "", "", "", "", ""), _
        Array("no_sec_v1_1", "No section on vlist1D: var 1", "", "", "This variable has no section, no subsection", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "19", "", "14", "78", "yes", "", "", "", "", "", ""), _
        Array("no_sec_v1_2", "No section on vlist1D: var 2", "", "", "This variable has no section, no subsection", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "20", "", "15", "60", "yes", "", "", "", "", "", ""), _
        Array("only_sec_v1_1", "Only section on vlist1D: var1", "", "", "This variable has a section, no subsection", "", "vlist1D-sheet1", "vlist1D", "Section only", "", "", "", "", "", "", "", "", "", "21", "", "16", "50", "yes", "", "", "", "", "", ""), _
        Array("only_sec_v1_2", "Only section on vlist1D: var2", "", "", "This variable has a section, no subsection", "", "vlist1D-sheet1", "vlist1D", "Section only", "", "", "", "", "", "", "", "", "", "22", "", "17", "66", "", "", "", "", "", "", ""), _
        Array("only_subsec_v1_1", "Only subsection on vlist1D: var1", "", "", "This variable has subsection, no section", "", "vlist1D-sheet1", "vlist1D", "", "Subsection only", "", "", "", "", "", "", "", "", "23", "", "18", "96", "yes", "", "", "", "", "", ""), _
        Array("only_subsec_v1_2", "Only subsection on vlist1D: var2", "", "", "This variable has a subsection, no section", "", "vlist1D-sheet1", "vlist1D", "", "Subsection only", "", "", "", "", "", "", "", "", "24", "", "19", "90", "yes", "", "", "", "", "", ""))
End Function

Private Function FixtureRowsChunk2() As Variant
    FixtureRowsChunk2 = Array( _
        Array("only_subsec_v1_3", "Only subsection on vlist1D: var3", "", "", "This variable has a subsection, no section", "", "vlist1D-sheet1", "vlist1D", "", "Subsection only", "", "", "", "", "", "", "", "", "25", "", "20", "90", "yes", "", "", "", "", "", ""), _
        Array("classif_v1", "Classification", "", "", "Classification variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "26", "", "21", "75", "yes", "", "", "", "", "", ""), _
        Array("regbook_v1", "Registered in book", "", "", "Registered", "", "vlist1D-sheet1", "vlist1D", "Register", "", "", "registered", "", "", "", "", "", "", "27", "", "22", "75", "yes", "", "", "", "", "", ""), _
        Array("perid_v1", "Personal identifier", "", "", "Used as identifier", "", "vlist1D-sheet1", "vlist1D", "Register", "", "", "", "personal", "", "", "", "", "", "28", "", "23", "71", "yes", "", "", "", "", "", ""), _
        Array("type_text_v1", "Type text", "", "", "Text variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "text", "", "", "", "", "", "29", "", "24", "71", "yes", "", "", "", "", "", ""), _
        Array("type_int_v1", "Type integer", "", "", "Integer variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "integer", "", "", "", "", "", "30", "", "25", "42", "yes", "", "", "", "", "", ""), _
        Array("type_dec_v1", "Type decimal", "", "", "Decimal variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "decimal", "", "", "", "", "", "31", "", "26", "32", "yes", "", "", "", "", "", ""), _
        Array("type_bool_v1", "Type boolean", "", "", "Boolean variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "boolean", "", "", "", "", "", "32", "", "27", "36", "yes", "", "", "", "", "", ""), _
        Array("type_date_v1", "Type date", "", "", "Date variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "date", "", "", "", "", "", "33", "", "28", "64", "yes", "", "", "", "", "", ""), _
        Array("type_time_v1", "Type time", "", "", "Time variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "time", "", "", "", "", "", "34", "", "29", "69", "yes", "", "", "", "", "", ""), _
        Array("type_datetime_v1", "Type datetime", "", "", "Datetime variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "datetime", "", "", "", "", "", "35", "", "30", "40", "yes", "", "", "", "", "", ""), _
        Array("type_duration_v1", "Type duration", "", "", "Duration variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "duration", "", "", "", "", "", "36", "", "31", "32", "yes", "", "", "", "", "", ""), _
        Array("type_percent_v1", "Type percentage", "", "", "Percentage variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "percentage", "", "", "", "", "", "37", "", "32", "60", "yes", "", "", "", "", "", ""), _
        Array("type_currency_v1", "Type currency", "", "", "Currency variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "currency", "", "", "", "", "", "38", "", "33", "74", "yes", "", "", "", "", "", ""), _
        Array("format_dec_v1", "Format decimal", "", "", "Format decimal variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "", "", "", "", "39", "", "34", "40", "yes", "", "", "", "", "", ""), _
        Array("format_date_v1", "Format date", "", "", "Format date variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "d-mmm-yyyy", "", "", "", "40", "", "35", "26", "yes", "", "", "", "", "", ""), _
        Array("format_time_v1", "Format time", "", "", "Format time variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm", "", "", "", "41", "", "36", "26", "yes", "", "", "", "", "", ""), _
        Array("format_datetime_v1", "Format datetime", "", "", "Format datetime variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "dd-mmm-yyyy hh:mm", "", "", "", "42", "", "37", "24", "yes", "", "", "", "", "", ""), _
        Array("format_duration_v1", "Format duration", "", "", "Format duration variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm", "", "", "", "43", "", "38", "94", "yes", "", "", "", "", "", ""), _
        Array("format_text_v1", "Format text", "", "", "Format text variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "", "", "", "", "44", "", "39", "49", "yes", "", "", "", "", "", ""))
End Function

Private Function FixtureRowsChunk3() As Variant
    FixtureRowsChunk3 = Array( _
        Array("format_currency_v1", "Format currency", "", "", "Format currency variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "€", "", "", "", "45", "", "40", "49", "yes", "", "", "", "", "", ""), _
        Array("format_percentage_v1", "Format percentage", "", "", "Format percentage variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "percentage", "", "", "", "46", "", "41", "28", "yes", "", "", "", "", "", ""), _
        Array("format_duration_v2", "Format duration v2", "", "", "Format duration variable v2", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm:ss", "", "", "", "47", "", "42", "28", "yes", "", "", "", "", "", ""), _
        Array("format_custom_v1", "Format custom", "", "", "Format custom variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "custom", "", "", "", "48", "", "43", "72", "yes", "", "", "", "", "", ""), _
        Array("note_v1", "Note", "", "", "Note on variable", "The note", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "49", "", "44", "72", "yes", "", "", "", "", "", ""), _
        Array("alert_v1", "Alert", "", "", "Alert variable", "Alert note", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "alert", "50", "", "45", "26", "yes", "", "", "", "", "", ""), _
        Array("message_v1", "Message", "", "", "Message variable", "Message note", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "message", "51", "", "46", "46", "yes", "", "", "", "", "", ""), _
        Array("format_condition_v1", "Format condition", "", "", "Format condition variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "52", "", "47", "46", "yes", "", "", "", "", "", ""), _
        Array("format_values_v1", "Format values", "", "", "Format values variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "53", "", "48", "68", "yes", "", "", "", "", "", ""), _
        Array("format_values_v2", "Format values v2", "", "", "Format values variable v2", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "54", "", "49", "68", "yes", "", "", "", "", "", ""), _
        Array("format_condition_v2", "Format condition v2", "", "", "Format condition variable v2", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "55", "", "50", "68", "yes", "", "", "", "", "", ""), _
        Array("format_values_v3", "Format values v3", "", "", "Format values variable v3", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "56", "", "51", "68", "yes", "", "", "", "", "", ""), _
        Array("format_condition_v3", "Format condition v3", "", "", "Format condition variable v3", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "format", "57", "", "52", "68", "yes", "", "", "", "", "", ""), _
        Array("list_auto_v1", "List auto", "", "", "list_auto variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "list_auto", "list_correct_order", "", "58", "", "53", "71", "yes", "", "", "", "", "", ""), _
        Array("formula_v1", "Formula", "", "", "Formula variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "formula", "IF(1=1, ""ok"", ""ko"")", "", "59", "", "54", "38", "yes", "", "", "", "", "", ""), _
        Array("formula_v2", "Formula v2", "", "", "Formula variable v2", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "formula", "IF(1=2, ""ok"", ""ko"")", "", "60", "", "55", "38", "yes", "", "", "", "", "", ""), _
        Array("formula_v3", "Formula v3", "", "", "Formula variable v3", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "formula", "IF(2=2, ""ok"", ""ko"")", "", "61", "", "56", "38", "yes", "", "", "", "", "", ""), _
        Array("formula_v4", "Formula v4", "", "", "Formula variable v4", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "formula", "IF(3=2, ""ok"", ""ko"")", "", "62", "", "57", "38", "yes", "", "", "", "", "", ""), _
        Array("geo_v1", "Geo variable", "", "", "Geo variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "63", "", "58", "69", "yes", "", "", "", "", "", ""), _
        Array("hf_v1", "HF variable", "", "", "HF variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "", "", "", "", "", "", "64", "", "59", "69", "yes", "", "", "", "", "", ""))
End Function

Private Function FixtureRowsChunk4() As Variant
    FixtureRowsChunk4 = Array( _
        Array("date_entry_h2", "Date on hlist2D", "", "", "TODAY()-365 < Date < TODAY()", "", "hlist2D-sheet1", "hlist2D", "Validation", "", "", "", "", "date", "", "", "", "", "66", "30", "", "58", "yes", "TODAY() - 365", "TODAY()", "error", "TODAY() - 365 < Date < TODAY()", "", ""), _
        Array("date_valid_h2_1", "Date validation on hlist2D: 1", "", "", "Date on hlist2D < Date < TODAY() + 365", "", "hlist2D-sheet1", "hlist2D", "Validation", "", "", "", "", "date", "", "", "", "", "67", "31", "", "61", "yes", "date_entry_h2", "TODAY() + 365", "warning", "Date on hlist2D < Date < TODAY() + 365", "", ""), _
        Array("date_valid_h2_2", "Date validation on hlist2D: 2", "", "", "TODAY()-365 < Date < TODAY()", "", "hlist2D-sheet1", "hlist2D", "Validation", "", "", "", "", "date", "d-mmm-yyyy", "", "", "", "68", "32", "", "93", "yes", "TODAY() - 365", "TODAY()", "info", "TODAY() - 365 < Date < TODAY()", "", ""), _
        Array("num_valid_h2", "Numeric validation on hlist2D", "", "", "Value > 0", "", "hlist2D-sheet1", "hlist2D", "Validation", "", "", "", "", "decimal", "", "", "", "", "69", "33", "", "58", "", "0", "", "info", "Value > 0", "", ""), _
        Array("int_valid_h2", "Integer validation on hlist2D", "", "", "Value < 0", "", "hlist2D-sheet1", "hlist2D", "Validation", "", "", "", "", "integer", "", "", "", "", "70", "34", "", "60", "", "", "0", "info", "Value < 0", "", ""), _
        Array("date_form_h2", "Variable formated as date", "", "", "Formated as dd-mmm-yyyy", "", "hlist2D-sheet1", "hlist2D", "Format", "", "", "", "", "date", "d-mmm-yyyy", "", "", "", "71", "35", "", "54", "yes", "", "", "", "", "", ""), _
        Array("curr_form_h2", "Variable formated as currency", "", "", "Formated as currency €", "", "hlist2D-sheet1", "hlist2D", "Format", "", "", "", "", "decimal", "euros", "", "", "", "72", "36", "", "81", "yes", "", "", "", "", "", ""), _
        Array("perc_form_h2", "Variable formated as percentage", "", "", "Formated as % ", "", "hlist2D-sheet1", "hlist2D", "Format", "", "", "", "", "decimal", "percentage3", "", "", "", "73", "37", "", "27", "yes", "", "", "", "", "", ""), _
        Array("vis_hidd_reg_h2", "Visible, but hidden in the register", "", "", "Hidden in the register", "", "hlist2D-sheet1", "hlist2D", "Register", "", "", "hidden", "", "", "", "", "", "", "74", "38", "", "99", "yes", "", "", "", "", "", ""), _
        Array("hid_end_h2", "Hidden variable at the end", "", "", "Should be hidden", "", "hlist2D-sheet1", "hlist2D", "", "", "hidden", "", "", "", "", "", "", "", "75", "39", "", "28", "yes", "", "", "", "", "", ""), _
        Array("lauto_drop_h2", "List auto dropdown from another sheet", "", "", "Populated from another sheet, choi_h2", "", "hlist2D-sheet2", "hlist2D", "", "Value OF", "", "", "", "", "", "list_auto", "choi_h2", "", "76", "", "", "6", "yes", "", "", "", "", "", ""), _
        Array("val_of_text_h2", "Value of a text variable", "", "", "Formula, match value of another sheet", "", "hlist2D-sheet2", "hlist2D", "", "Value OF", "", "", "", "", "", "formula", "VALUE_OF(lauto_drop_h2, choi_h2, text_h2)", "", "77", "", "", "2", "yes", "", "", "", "", "", ""), _
        Array("val_of_int_h2", "Value of integer variable", "", "", "Formula, match value of another sheet", "", "hlist2D-sheet2", "hlist2D", "", "Value OF", "", "", "", "", "", "formula", "VALUE_OF(lauto_drop_h2, choi_h2, int_valid_h2)", "", "78", "", "", "55", "yes", "", "", "", "", "", ""), _
        Array("val_of_dec_h2", "Value of decimal variable", "", "", "Formula, match value of another sheet", "", "hlist2D-sheet2", "hlist2D", "", "Value OF", "", "", "", "", "", "formula", "VALUE_OF(lauto_drop_h2, choi_h2, num_valid_h2)", "", "79", "", "", "21", "yes", "", "", "", "", "", ""), _
        Array("val_of_date_h2", "Value of date variable", "", "", "Formula, match value of another sheet", "", "hlist2D-sheet2", "hlist2D", "", "Value OF", "", "", "", "", "", "formula", "VALUE_OF(lauto_drop_h2, choi_h2, date_form_h2)", "", "80", "", "", "58", "yes", "", "", "", "", "", ""), _
        Array("cond_test_h1", "Test on conditonal formatting", "", "", "Formula, should be hidden", "", "hlist2D-sheet1", "hlist2D", "", "Conditonal Formatting", "hidden", "hidden", "", "", "", "formula", "IF(choi_h2 = ""A"", 1, 0)", "", "81", "45", "", "20", "yes", "", "", "", "", "", ""), _
        Array("cond_val_h1", "Value on conditional formatting", "", "", "Test for conditonal formatting, should be in gray", "", "hlist2D-sheet1", "hlist2D", "", "Conditonal Formatting", "", "", "", "", "", "", "", "", "82", "46", "", "15", "yes", "", "", "", "", "cond_test_h1", ""))
End Function
