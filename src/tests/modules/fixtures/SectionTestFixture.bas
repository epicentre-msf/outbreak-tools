Attribute VB_Name = "SectionTestFixture"

Option Explicit
Option Private Module

'@Folder("Tests")
'@ModuleDescription("Shared helpers for section writer tests")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const HEADER_ROW As Long = 1
Private Const DATA_START_ROW As Long = 2

Public Function SectionFixtureHeaders() As Variant
    SectionFixtureHeaders = Array( _
        "main section", _
        "sub section", _
        "sheet name", _
        "variable name", _
        "column index", _
        "crf index", _
        "main label", _
        "sub label", _
        "variable type", _
        "variable format", _
        "status", _
        "note", _
        "control", _
        "min", _
        "max", _
        "alert", _
        "message")
End Function

Public Function HorizontalSectionRows() As Variant
    HorizontalSectionRows = Array( _
        Array("Section H", "Sub H1", "HSection", "var_h1", 4, 12, "Main H1", "Sub H1", "text", vbNullString, "active", "", "text", vbNullString, vbNullString, "warning", ""), _
        Array("Section H", "Sub H1", "HSection", "var_h2", 6, 12, "Main H2", "Sub H2", "text", vbNullString, "active", "", "text", vbNullString, vbNullString, "warning", ""), _
        Array("Section H", "Sub H2", "HSection", "var_h3", 8, 16, "Main H3", "Sub H3", "text", vbNullString, "active", "", "text", vbNullString, vbNullString, "warning", ""))
End Function

Public Function VerticalSectionRows() As Variant
    VerticalSectionRows = Array( _
        Array("Section V", "Sub V1", "VSection", "var_v1", 10, 0, "Main V1", "Sub V1", "text", vbNullString, "active", "", "text", vbNullString, vbNullString, "warning", ""), _
        Array("Section V", "Sub V1", "VSection", "var_v2", 12, 0, "Main V2", "Sub V2", "text", vbNullString, "active", "", "text", vbNullString, vbNullString, "warning", ""))
End Function

Public Function CreateSectionContext(ByVal baseName As String, _
                                     ByVal rows As Variant, _
                                     ByVal startRow As Long, _
                                     ByVal design As ILLFormat) As ILLSectionContext

    Dim dataSheetName As String
    Dim dataSheet As Worksheet
    Dim headers As Variant
    Dim rowCount As Long
    Dim columnCount As Long
    Dim idx As Long
    Dim dictStub As LLSectionDictionaryStub
    Dim specsStub As LLVarContextSpecsStub
    Dim linelistStub As LLVarContextLinelistStub
    Dim uniqueSheets As Collection
    Dim sheetName As String
    Dim rowValues As Variant

    dataSheetName = baseName & "_Data"
    headers = SectionFixtureHeaders()
    rowCount = UBound(rows) - LBound(rows) + 1
    columnCount = UBound(headers) - LBound(headers) + 1

    Set dataSheet = TestHelpers.EnsureWorksheet(dataSheetName)
    dataSheet.Cells.Clear
    TestHelpers.WriteRow dataSheet.Cells(HEADER_ROW, 1), headers
    TestHelpers.WriteMatrix dataSheet.Cells(DATA_START_ROW, 1), TestHelpers.RowsToMatrix(rows)

    Set uniqueSheets = New Collection
    On Error Resume Next
    For idx = LBound(rows) To UBound(rows)
        rowValues = rows(idx)
        sheetName = CStr(rowValues(2))
        uniqueSheets.Add sheetName, sheetName
    Next idx
    On Error GoTo 0

    For idx = 1 To uniqueSheets.Count
        sheetName = uniqueSheets(idx)
        TestHelpers.EnsureWorksheet sheetName
    Next idx

    Set dictStub = New LLSectionDictionaryStub
    dictStub.Configure dataSheet, DATA_START_ROW, 1, DATA_START_ROW + rowCount - 1, columnCount

    Set specsStub = New LLVarContextSpecsStub
    specsStub.SetDesignFormat design

    Set linelistStub = New LLVarContextLinelistStub
    linelistStub.UseWorkbook ThisWorkbook
    linelistStub.UseSpecs specsStub
    linelistStub.UseDictionary dictStub

    Dim context As ILLSectionContext
    Set context = New LLSectionContext
    context.Initialise linelistStub, startRow

    Set CreateSectionContext = context
End Function

