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
    FixtureHeadersArray = Split("Variable Name|Main Label|Dev Comments|Editable Label|Sub Label|Note|Sheet Name|Sheet Type|Main Section|Sub Section|Status|Register Book|Personal Identifier|Variable Type|Variable Format|Control|Control Details|Unique|Export 1|Export 2|Export 3|Export 4|Export 5|Min|Max|Alert|Message|Formatting Condition|Formatting Values|Lock Cells|Column Index", "|")
End Function

Private Function FixtureRowsArray() As Variant
    FixtureRowsArray = CombineRowSets( _
        FixtureRowsChunk1(), _
        FixtureRowsChunk2(), _
        FixtureRowsChunk3(), _
        FixtureRowsChunk4(), _
        FixtureRowsChunk5(), _
        FixtureRowsChunk6() _
    )
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
    Dim sheetAssignments As Collection
    Dim sheetIndex As Long
    Dim tableIndex As Long

    sheetIndex = HeaderIndexOf(headers, SHEET_NAME_HEADER)
    tableIndex = HeaderIndexOf(headers, TABLE_NAME_HEADER)

    ReDim result(LBound(rows) To UBound(rows))

    Set sheetAssignments = New Collection

    For idx = LBound(rows) To UBound(rows)
        result(idx) = InsertTableNameValue(rows(idx), sheetIndex, tableIndex, sheetAssignments)
    Next idx

    AddTableNamesToRows = result
End Function

Private Function InsertTableNameValue( ByVal rowValues As Variant, _
                                       ByVal sheetIndex As Long, _
                                       ByVal tableIndex As Long, _
                                       ByVal sheetAssignments As Collection) As Variant
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

Private Function ResolveTableName(ByVal sheetName As String, ByVal assignments As Collection) As String
    Dim sheetKey As String
    Dim existing As Variant
    Dim errNumber As Long
    Dim tableName As String

    sheetKey = LCase$(sheetName)
    If Len(sheetKey) = 0 Then sheetKey = "<empty>"

    On Error Resume Next
    existing = assignments(sheetKey)
    errNumber = Err.Number
    On Error GoTo 0

    If errNumber <> 0 Then
        tableName = "table" & CStr(assignments.Count + 1)
        assignments.Add tableName, sheetKey
    Else
        tableName = CStr(existing)
    End If

    ResolveTableName = tableName
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
        Array("hid_beg_v1","Hidden variable at the begining","","","Should be hidden","","vlist1D-sheet1","vlist1D","Hidden Section","","hidden","","","","","","","","6","","1","83","yes","","","","","","",""), _
        Array("choi_v1","Choices on vlist1D","","","Values: A, B, C","","vlist1D-sheet1","vlist1D","Controls","","","","","","","choice_manual","list_correct_order","","7","","2","5","","","","","","","",""), _
        Array("choi_mult_v1","Choice multiple on vlist1D","","","Multiple values: A, B, C","","vlist1D-sheet1","vlist1D","Controls","","","","","","","choice_multiple","list_multiple","","8","","3","36","","","","","","","",""), _
        Array("choi_ord_v1","Choice order on vlist1D","","","Values: B, C, A","","vlist1D-sheet1","vlist1D","Controls","","","","","","","choice_manual","list_uncorrect_order","","9","","4","94","yes","","","","","","",""), _
        Array("choi_cust_v1","Custom choices on vlist1D","","","Random input by the user","","vlist1D-sheet1","vlist1D","Controls","","","","","","","choice_custom","","","10","","5","57","yes","","","","","","",""), _
        Array("form_v1","Formula on vlist1D","","","Formula","","vlist1D-sheet1","vlist1D","Controls","","","","","","","formula","IF(ISBLANK(choi_v1), """", choi_v1 & ""-OK"")","","11","","6","8","","","","","","","",""), _
        Array("brok_form_v1","Broken formula on vlist1D","","","This formula should fail","","vlist1D-sheet1","vlist1D","Controls","","","","","","","formula","IF(ISBLANK(choi_v2), """", choi_v2 & ""-OK"")","","12","","7","94","yes","","","","","","",""), _
        Array("hid_v1","Hidden variable in the middle","","","Should be hidden","","vlist1D-sheet1","vlist1D","Status","","hidden","","","","","","","","13","","8","36","yes","","","","","","",""), _
        Array("opt_hid_v1","Optional hidden variable on vlist1D","","","Should be hidden, the user can unhide","","vlist1D-sheet1","vlist1D","Status","","optional, hidden","","","","","","","","14","","9","43","","","","","","","",""), _
        Array("opt_vis_v1","Optional visible variable on vlist1D","","","Should be visible, the user can hide","","vlist1D-sheet1","vlist1D","Status","","optional, visible","","","","","","","","15","","10","80","yes","","","","","","",""), _
        Array("mand_v1","Mandatory variable on vlist1D","","","Mandatory","","vlist1D-sheet1","vlist1D","Status","","mandatory","","","","","","","","16","","11","65","yes","","","","","","",""), _
        Array("c_wh_v1","Case when on vlist1D","","","Case When formula","","vlist1D-sheet1","vlist1D","Controls","","","","","","","case_when","CASE_WHEN(choi_ord_v1 = ""A"", ""Choice order is A"", choi_ord_v1 = ""B"", ""Choice order is B"", ""Unknown choice order"")","","17","","12","22","yes","","","","","","",""), _
        Array("choi_form_v1","Choice formula on vlist1D","","","Choice Formula","","vlist1D-sheet1","vlist1D","Controls","","","","","","","choice_formula","CHOICE_FORMULA(list_multiple, choi_v1 = ""A"", ""choice 1"", choi_v1 = ""B"", ""choice 2"", choi_v1 = ""C"", c_wh_v1)","","18","","13","60","yes","","","","","","",""), _
        Array("no_sec_v1_1","No section on vlist1D: var 1","","","This variable has no section, no subsection","","vlist1D-sheet1","vlist1D","","","","","","","","","","","19","","14","48","yes","","","","","","",""), _
        Array("no_sec_v1_2","No section on vlist1D: var 2","","","This variable has no section, no subsection","","vlist1D-sheet1","vlist1D","","","","","","","","","","","20","","15","83","","","","","","","","") _
    )
End Function

Private Function FixtureRowsChunk2() As Variant
    FixtureRowsChunk2 = Array( _
        Array("only_sec_v1_1","Only section on vlist1D: var1","","","This variable has a section, no subsection","","vlist1D-sheet1","vlist1D","Section only","","","","","","","","","","21","","16","13","yes","","","","","","",""), _
        Array("only_sec_v1_2","Only section on vlist1D: var2","","","This variable has a section, no subsection","","vlist1D-sheet1","vlist1D","Section only","","","","","","","","","","22","","17","42","","","","","","","",""), _
        Array("only_subsec_v1_1","Only subsection on vlist1D: var1","","","This variable has subsection, no section","","vlist1D-sheet1","vlist1D","","Subsection only","","","","","","","","","23","","18","43","yes","","","","","","",""), _
        Array("only_subsec_v1_2","Only subsection on vlist1D: var2","","","This variable has a subsection, no section","","vlist1D-sheet1","vlist1D","","Subsection only","","","","","","","","","24","","19","38","yes","","","","","","",""), _
        Array("date_v1","Date on vlist1D","","","TODAY()-365 < Date < TODAY()","","vlist1D-sheet1","vlist1D","Validation","Date validation","","","","date","","","","","25","","20","3","yes","TODAY() - 365","TODAY()","error","TODAY() - 365 < Date < TODAY()","","",""), _
        Array("int_v1","Integer on vlist1D","","","Value > 0","","vlist1D-sheet1","vlist1D","Validation","","","","","integer","","","","","26","","21","3","yes","0","","warning","Shoud be > 0","","",""), _
        Array("dec2_v1","Decimal 2 digits on vlist1D","","","Value < 1","","vlist1D-sheet1","vlist1D","Validation","","","","","decimal","","","","","27","","22","73","","","1","info","Should be < 1","","",""), _
        Array("date_vali_v1","Date validation on vlist1D","","","date_v1 < Date < TODAY() + 365","This is a test for a note","vlist1D-sheet1","vlist1D","Validation","Date validation","","","","date","","","","","28","","23","68","yes","date_v1","TODAY() + 365","warning","Date on vlist1D < Date < TODAY() + 365","","",""), _
        Array("num_vali_v1","Numeric validation on vlist1D","","","Should be between [3, 10]","","vlist1D-sheet1","vlist1D","Validation","","","","","","","","","","29","","24","1","","3","10","error","Should be between [3, 10]","","",""), _
        Array("exp_var_v1","Variable used in export vlist1D","","","Use for export name","","vlist1D-sheet1","vlist1D","","","","","","text","","","","","30","","25","32","yes","","","","","","",""), _
        Array("perc_form_v1","Variable formated as percentage","","","Formated as %","","vlist1D-sheet1","vlist1D","Format","","","","","decimal","percentage2","","","","31","","26","13","yes","","","","","","",""), _
        Array("num_form_v1","Variable can be rounded to 4 digits","","","Can be rounded to 4 digits","","vlist1D-sheet1","vlist1D","Format","","","","","decimal","round4","","","","32","","27","36","yes","","","","","","",""), _
        Array("date_form_v1","Variable formated as date","","","Formated as dd-mmm-yyyy","","vlist1D-sheet1","vlist1D","Format","","","","","","d-mmm-yyyy","","","","33","","28","82","","","","","","","",""), _
        Array("curr_form_v1","Variable formated as currency","","","Formated as currency $","","vlist1D-sheet1","vlist1D","Format","","","","","","dollars","","","","34","","29","8","","","","","","","",""), _
        Array("ed_var_v1","Editable variable label vlist1D","","yes","You can change the label","","vlist1D-sheet1","vlist1D","","","","","","","","","","","35","","30","16","","","","","","","","") _
    )
End Function

Private Function FixtureRowsChunk3() As Variant
    FixtureRowsChunk3 = Array( _
        Array("hid_end_v1","Hidden variable at the end","","","Should be hidden, personal identifier on v1","","vlist1D-sheet1","vlist1D","","","","","yes","","","","","","36","","31","95","yes","","","","","","",""), _
        Array("hid_beg_h2","Hidden variable at the begining","","","Should be hidden","","hlist2D-sheet1","hlist2D","","","","","","","","","","","37","1","","82","","","","","","","",""), _
        Array("pers_id_h2","Personal identifier variable","","","Personal identifier","","hlist2D-sheet1","hlist2D","","","","","","","","","","","38","2","","32","yes","","","","","","",""), _
        Array("uni_h2","Unique variable","","","Unique variable, duplicates in red","","hlist2D-sheet1","hlist2D","","","","","","","","","","yes","39","3","","19","yes","","","","","","",""), _
        Array("bad name h2","Bad variable name, shoud be renamed","","","Bad variable name which is corrected automatically","","hlist2D-sheet1","hlist2D","","","","","","","","","","","40","4","","35","yes","","","","","","",""), _
        Array("date_h2","Date on hlist2D","","","","","hlist2D-sheet1","hlist2D","Var Type","","","","","date","","","","","41","5","","46","","","","","","","",""), _
        Array("int_h2","Integer on hlist2D","","","","","hlist2D-sheet1","hlist2D","Var Type","","","","","integer","","","","","42","6","","83","yes","","","","","","",""), _
        Array("text_h2","Random text variable","","yes","You can change the label","","hlist2D-sheet1","hlist2D","","","","","","","","","","","43","7","","45","yes","","","","","","",""), _
        Array("dec2_h2","Decimal 2 digits on hlist2D","","","","This is a test for a note","hlist2D-sheet1","hlist2D","Format","","","","","decimal","round2","","","","44","8","","90","yes","","","","","","",""), _
        Array("no_sec_h2_1","No section : var1","","","This variable has no section, no subsection","","hlist2D-sheet1","hlist2D","","","","","","","","","","","45","9","","30","yes","","","","","","",""), _
        Array("no_sec_h2_2","No section: var2","","","This variable has no section, no subsection","","hlist2D-sheet1","hlist2D","","","","","","","","","","","46","10","","19","yes","","","","","","",""), _
        Array("only_sec_h2_1","Only sec: var1","","","This variable has a section, no subsection","","hlist2D-sheet1","hlist2D","Section only","","","","","","","","","","47","11","","41","yes","","","","","","",""), _
        Array("only_sec_h2_2","Only sec: var2","","","This variable has a section, no subsection","","hlist2D-sheet1","hlist2D","Section only","","","","","","","","","","48","12","","62","yes","","","","","","",""), _
        Array("only_subsec_h2_1","Only subsec: var1","","","This variable has a subsection, no section","","hlist2D-sheet1","hlist2D","","Subsection only","","","","","","","","","49","13","","48","yes","","","","","","",""), _
        Array("only_subsec_h2_2","Only subsec: var2","","","This variable has a subsection, no section","","hlist2D-sheet1","hlist2D","","Subsection only","","","","","","","","","50","14","","31","yes","","","","","","","") _
    )
End Function

Private Function FixtureRowsChunk4() As Variant
    FixtureRowsChunk4 = Array( _
        Array("mand_h2","Mandatory variable on hlist2D","","","Mandatory","","hlist2D-sheet1","hlist2D","Status","","mandatory","","","","","","","","51","15","","58","","","","","","","",""), _
        Array("hid_h2","Hidden variable in the middle","","","Should be hidden","","hlist2D-sheet1","hlist2D","Status","","hidden","","","","","","","","52","16","","1","yes","","","","","","",""), _
        Array("opt_hid_h2","Optional hidden variable on hlist2D","","","Should be hidden, the user can unhide","","hlist2D-sheet1","hlist2D","Status","","optional, hidden","","","","","","","","53","17","","70","","","","","","","",""), _
        Array("opt_vis_h2","Optional visible variable on hlist2D","","","Should be visible, the user can hide","","hlist2D-sheet1","hlist2D","","","optional, visible","","","","","","","","54","18","","46","","","","","","","",""), _
        Array("choi_h2","Choice on hlist2D","","","Values: A, B, C","","hlist2D-sheet1","hlist2D","Controls","","","","","","","choice_manual","list_correct_order","","55","19","","10","yes","","","","","","",""), _
        Array("choi_mult_h2","Choice multiple on hlist2D","","","Multiple values: A, B, C","","hlist2D-sheet1","hlist2D","Controls","","","","","","","choice_multiple","list_multiple","","56","20","","3","yes","","","","","","",""), _
        Array("choi_cust_h2","Custom choices on hlist2D","","","Random input by the user","","hlist2D-sheet1","hlist2D","Controls","","","","","","","choice_custom","","","57","21","","69","","","","","","","",""), _
        Array("form_h2","Formula on hlist2D","","","Formula","","hlist2D-sheet1","hlist2D","Controls","","","","","","","formula","IF(ISBLANK(choi_h2), """", choi_h2 & ""-OK"")","","58","22","","8","yes","","","","","","",""), _
        Array("brok_form_h2","Broken formula on hlist2D","","","This formula should fail","","hlist2D-sheet1","hlist2D","Controls","","","","","","","formula","IF(ISBLANK(choi_h2), """", choi_h2 & + ""-OK"")","","59","23","","92","","","","","","","",""), _
        Array("lauto_man_h2","List auto variable manual","","","List auto populated from ""Ramdom text variable""","","hlist2D-sheet1","hlist2D","Controls","","","","","","","list_auto","text_h2","","60","24","","35","yes","","","","","","",""), _
        Array("choi_form_h2","Choice formula on hlist2D","","","Choice Formula","","hlist2D-sheet1","hlist2D","Controls","","","","","","","choice_formula","CHOICE_FORMULA(list_multiple, choi_h2 = ""A"", ""choice 1"", choi_h2 = ""B"", ""choice 2"", choi_h2 = ""C"", c_wh_h2)","","61","25","","50","yes","","","","","","",""), _
        Array("c_wh_h2","Case when on hlist2D","","","Case When formula","","hlist2D-sheet1","hlist2D","Controls","","","","","","","case_when","CASE_WHEN(choi_ord_v1 = ""A"", ""Choice order is A"", choi_ord_v1 = ""B"", ""Choice order is B"", ""Unknown choice order"")","","62","26","","20","yes","","","","","","",""), _
        Array("geo_h2","Geo on hlist2D","","","Residence","","hlist2D-sheet1","hlist2D","Controls","","","","","","","geo","","","63","27","","21","yes","","","","","","",""), _
        Array("hf_h2","HF on hlist2D","","","Health facility","","hlist2D-sheet1","hlist2D","Controls","","","","","","","hf","","","64","28","","1","yes","","","","","","",""), _
        Array("lauto_comp_h2","List auto variable from formula","","","List auto populated from ""Formula on hlist2D""","","hlist2D-sheet1","hlist2D","Controls","","","","","","","list_auto","form_h2","","65","29","","13","","","","","","","","") _
    )
End Function

Private Function FixtureRowsChunk5() As Variant
    FixtureRowsChunk5 = Array( _
        Array("date_entry_h2","Date on hlist2D","","","TODAY()-365 < Date < TODAY()","","hlist2D-sheet1","hlist2D","Validation","","","","","date","","","","","66","30","","39","yes","TODAY() - 365","TODAY()","error","TODAY() - 365 < Date < TODAY()","","",""), _
        Array("date_valid_h2_1","Date validation on hlist2D: 1","","","Date on hlist2D < Date < TODAY() + 365","","hlist2D-sheet1","hlist2D","Validation","","","","","date","","","","","67","31","","60","yes","date_entry_h2","TODAY() + 365","warning","Date on hlist2D < Date < TODAY() + 365","","",""), _
        Array("date_valid_h2_2","Date validation on hlist2D: 2","","","TODAY()-365 < Date < TODAY()","","hlist2D-sheet1","hlist2D","Validation","","","","","date","d-mmm-yyyy","","","","68","32","","58","yes","TODAY() - 365","TODAY()","info","TODAY() - 365 < Date < TODAY()","","",""), _
        Array("num_valid_h2","Numeric validation on hlist2D","","","Value > 0","","hlist2D-sheet1","hlist2D","Validation","","","","","decimal","","","","","69","33","","40","yes","0","","info","Value > 0","","",""), _
        Array("int_valid_h2","Integer validation on hlist2D","","","Value < 0","","hlist2D-sheet1","hlist2D","Validation","","","","","integer","","","","","70","34","","53","yes","","0","info","Value < 0","","",""), _
        Array("date_form_h2","Variable formated as date","","","Formated as dd-mmm-yyyy","","hlist2D-sheet1","hlist2D","Format","","","","","date","d-mmm-yyyy","","","","71","35","","71","","","","","","","",""), _
        Array("curr_form_h2","Variable formated as currency","","","Formated as currency €","","hlist2D-sheet1","hlist2D","Format","","","","","decimal","euros","","","","72","36","","93","yes","","","","","","",""), _
        Array("perc_form_h2","Variable formated as percentage","","","Formated as % ","","hlist2D-sheet1","hlist2D","Format","","","","","decimal","percentage3","","","","73","37","","76","yes","","","","","","",""), _
        Array("vis_hidd_reg_h2","Visible, but hidden in the register","","","Hidden in the register","","hlist2D-sheet1","hlist2D","Register","","","hidden","","","","","","","74","38","","17","yes","","","","","","",""), _
        Array("hid_end_h2","Hidden variable at the end","","","Should be hidden","","hlist2D-sheet1","hlist2D","","","hidden","","","","","","","","75","39","","81","","","","","","","",""), _
        Array("lauto_drop_h2","List auto dropdown from another sheet","","","Populated from another sheet, choi_h2","","hlist2D-sheet2","hlist2D","","Value OF","","","","","","list_auto","choi_h2","","76","","","32","yes","","","","","","",""), _
        Array("val_of_text_h2","Value of a text variable","","","Formula, match value of another sheet","","hlist2D-sheet2","hlist2D","","Value OF","","","","","","formula","VALUE_OF(lauto_drop_h2, choi_h2, text_h2)","","77","","","14","yes","","","","","","",""), _
        Array("val_of_int_h2","Value of integer variable","","","Formula, match value of another sheet","","hlist2D-sheet2","hlist2D","","Value OF","","","","","","formula","VALUE_OF(lauto_drop_h2, choi_h2, int_valid_h2)","","78","","","72","yes","","","","","","",""), _
        Array("val_of_dec_h2","Value of decimal variable","","","Formula, match value of another sheet","","hlist2D-sheet2","hlist2D","","Value OF","","","","","","formula","VALUE_OF(lauto_drop_h2, choi_h2, num_valid_h2)","","79","","","79","yes","","","","","","",""), _
        Array("val_of_date_h2","Value of date variable","","","Formula, match value of another sheet","","hlist2D-sheet2","hlist2D","","Value OF","","","","","","formula","VALUE_OF(lauto_drop_h2, choi_h2, date_form_h2)","","80","","","42","yes","","","","","","","") _
    )
End Function

Private Function FixtureRowsChunk6() As Variant
    FixtureRowsChunk6 = Array( _ 
        Array("type_bool_v1", "Type boolean", "", "", "Boolean variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "boolean", "", "", "", "", "", "32", "", "27", "36", "yes", "", "", "", "", "", "", "39"), _
        Array("type_date_v1", "Type date", "", "", "Date variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "date", "", "", "", "", "", "33", "", "28", "64", "yes", "", "", "", "", "", "", "40"), _
        Array("type_time_v1", "Type time", "", "", "Time variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "time", "", "", "", "", "", "34", "", "29", "69", "yes", "", "", "", "", "", "", "41"), _
        Array("type_datetime_v1", "Type datetime", "", "", "Datetime variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "datetime", "", "", "", "", "", "35", "", "30", "40", "yes", "", "", "", "", "", "", "42"), _
        Array("type_duration_v1", "Type duration", "", "", "Duration variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "duration", "", "", "", "", "", "36", "", "31", "32", "yes", "", "", "", "", "", "", "43"), _
        Array("type_percent_v1", "Type percentage", "", "", "Percentage variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "percentage", "", "", "", "", "", "37", "", "32", "60", "yes", "", "", "", "", "", "", "44"), _
        Array("type_currency_v1", "Type currency", "", "", "Currency variable", "", "vlist1D-sheet1", "vlist1D", "", "", "", "", "currency", "", "", "", "", "", "38", "", "33", "74", "yes", "", "", "", "", "", "", "45"), _
        Array("format_dec_v1", "Format decimal", "", "", "Format decimal variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "", "", "", "", "39", "", "34", "40", "yes", "", "", "", "", "", "", "46"), _
        Array("format_date_v1", "Format date", "", "", "Format date variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "d-mmm-yyyy", "", "", "", "40", "", "35", "26", "yes", "", "", "", "", "", "", "47"), _
        Array("format_time_v1", "Format time", "", "", "Format time variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm", "", "", "", "41", "", "36", "26", "yes", "", "", "", "", "", "", "48"), _
        Array("format_datetime_v1", "Format datetime", "", "", "Format datetime variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "dd-mmm-yyyy hh:mm", "", "", "", "42", "", "37", "24", "yes", "", "", "", "", "", "", "49"), _
        Array("format_duration_v1", "Format duration", "", "", "Format duration variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm", "", "", "", "43", "", "38", "94", "yes", "", "", "", "", "", "", "50"), _
        Array("format_text_v1", "Format text", "", "", "Format text variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "", "", "", "", "44", "", "39", "49", "yes", "", "", "", "", "", "", "51"), _
        Array("format_currency_v1", "Format currency", "", "", "Format currency variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "€", "", "", "", "45", "", "40", "49", "yes", "", "", "", "", "", "", "52"), _
        Array("format_percentage_v1", "Format percentage", "", "", "Format percentage variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "percentage", "", "", "", "46", "", "41", "28", "yes", "", "", "", "", "", "", "53"), _
        Array("format_duration_v2", "Format duration v2", "", "", "Format duration variable v2", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "hh:mm:ss", "", "", "", "47", "", "42", "28", "yes", "", "", "", "", "", "", "54"), _
        Array("format_custom_v1", "Format custom", "", "", "Format custom variable", "", "vlist1D-sheet1", "vlist1D", "", "Format", "", "", "", "", "custom", "", "", "", "48", "", "43", "72", "yes", "", "", "", "", "", "", "55"), _
        Array("cond_test_h1","Test on conditonal formatting","","","Formula, should be hidden","","hlist2D-sheet1","hlist2D","","Conditonal Formatting","hidden","hidden","","","","formula","IF(choi_h2 = ""A"", 1, 0)","","81","45","","77","yes","","","","","","",""), _
        Array("cond_val_h1","Value on conditional formatting","","","Test for conditonal formatting, should be in gray","","hlist2D-sheet1","hlist2D","","Conditonal Formatting","","","","","","","","","82","46","","30","yes","","","","","cond_test_h1","","") _
    )
End Function
