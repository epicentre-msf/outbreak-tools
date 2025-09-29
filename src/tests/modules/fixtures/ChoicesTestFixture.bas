Attribute VB_Name = "ChoicesTestFixture"
Attribute VB_Description = "Shared choices fixture for LLChoices tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Shared choices fixture for LLChoices tests")

'@section Fixture Cache
'===============================================================================

Private fixtureHeaders As Variant
Private fixtureRows As Variant

Private Sub EnsureFixtureLoaded()
    If IsEmpty(fixtureHeaders) Then fixtureHeaders = FixtureHeadersArray()
    If IsEmpty(fixtureRows) Then fixtureRows = FixtureRowsArray()
End Sub

'@section Worksheet Preparation
'===============================================================================

'@description Seed a worksheet with the choices fixture data.
'@param sheetName String. Worksheet to create or reset.
'@param targetBook Workbook hosting the worksheet.
Public Sub PrepareChoicesFixture(ByVal sheetName As String, _
                                 Optional ByVal targetBook As Workbook)
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    BusyApp
    
    EnsureFixtureLoaded

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    Set sh = EnsureWorksheet(sheetName, wb)

    headerMatrix = RowsToMatrix(Array(fixtureHeaders))
    WriteMatrix sh.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(fixtureRows)
    WriteMatrix sh.Cells(2, 1), dataMatrix

End Sub

'@section Fixture Metadata
'===============================================================================

'@description Retrieve distinct list names from the fixture.
'@return Variant 1-D array of list names.
Public Function ChoicesFixtureDistinctLists() As Variant
    Dim names As Collection
    Dim rowData As Variant
    Dim value As String
    Dim result() As String
    Dim idx As Long

    EnsureFixtureLoaded

    Set names = New Collection

    On Error Resume Next
        For Each rowData In fixtureRows
            value = CStr(rowData(LBound(fixtureHeaders)))
            names.Add value, CStr(value)
        Next rowData
    On Error GoTo 0

    ReDim result(0 To names.Count - 1)
    For idx = 1 To names.Count
        result(idx - 1) = names(idx)
    Next idx

    ChoicesFixtureDistinctLists = result
End Function

'@description Get the number of data rows in the fixture.
'@return Long count of rows.
Public Function ChoicesFixtureRowCount() As Long
    EnsureFixtureLoaded
    ChoicesFixtureRowCount = UBound(fixtureRows) - LBound(fixtureRows) + 1
End Function

'@description Retrieve the fixture headers.
'@return Variant array of headers.
Public Function ChoicesFixtureHeaders() As Variant
    EnsureFixtureLoaded
    ChoicesFixtureHeaders = fixtureHeaders
End Function


'@section Fixture Data
'===============================================================================

Private Function FixtureHeadersArray() As Variant
    FixtureHeadersArray = Array("list name", "ordering list", "label", "short label")
End Function

Private Function FixtureRowsArray() As Variant
    FixtureRowsArray = Array( _
        Array("list_correct_order", 1, "A", "A short"), _
        Array("list_correct_order", 2, "B", "B short"), _
        Array("list_correct_order", 3, "C", "C short"), _
        Array("list_uncorrect_order", 3, "A", "A short"), _
        Array("list_uncorrect_order", 1, "B", vbNullString), _
        Array("list_uncorrect_order", 2, "C", "C short"), _
        Array("list_multiple", 1, "choice 1", "c1"), _
        Array("list_multiple", 2, "choice 2", "c2"), _
        Array("list_multiple", 3, "choice 3", vbNullString), _
        Array("list_multiple", 4, "choice 4", "c4"))
End Function
