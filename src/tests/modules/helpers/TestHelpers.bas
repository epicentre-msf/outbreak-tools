Attribute VB_Name = "TestHelpers"
Attribute VB_Description = "Utility helpers shared across tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@Folder("Tests")
'@ModuleDescription("Utility helpers shared across tests")

'@section Application State
'===============================================================================

'@description Suspend heavy Excel UI features while tests manipulate workbooks.
Public Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'@description Restore the Excel UI to its default behaviour after BusyApp.
Public Sub RestoreApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

'@section Workbooks
'===============================================================================

'@description Create a new workbook ready for test usage.
'@return Workbook freshly created.
Public Function NewWorkbook() As Workbook
    BusyApp
    Set NewWorkbook = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
End Function

'@description Close and discard a workbook if it exists.
'@param wb Workbook or Object reference to close.
Public Sub DeleteWorkbook(ByVal wb As Workbook)
    On Error Resume Next
        BusyApp
        wb.Close saveChanges:=False
    On Error GoTo 0
End Sub

'@section Worksheets
'===============================================================================

'@description Ensure a worksheet exists and is cleared.
'@param sheetName String. Name of the worksheet to create/reset.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Worksheet ensured for use.
Public Function EnsureWorksheet(ByVal sheetName As String, _
                                Optional ByVal targetBook As Workbook) As Worksheet

    Dim wb As Workbook
    Dim sh As Worksheet

    BusyApp

    If (targetBook Is Nothing) Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add
        sh.Name = sheetName
    End If

    ClearWorksheet sh
    Set EnsureWorksheet = sh
End Function

'@description Create a worksheet when missing and clear its cells.
'@param sheetName String. Name of the worksheet to reset.
Public Sub NewWorksheet(ByVal sheetName As String)
    Call EnsureWorksheet(sheetName)
End Sub

'@description Delete a worksheet if it exists.
'@param sheetName String. Worksheet to delete.
Public Sub DeleteWorksheet(ByVal sheetName As String)
    On Error Resume Next
        BusyApp
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

'@description Delete several worksheets in a single call.
'@param sheetNames ParamArray list of worksheet names.
Public Sub DeleteWorksheets(ParamArray sheetNames() As Variant)
    Dim idx As Long

    For idx = LBound(sheetNames) To UBound(sheetNames)
        DeleteWorksheet CStr(sheetNames(idx))
    Next idx
End Sub

'@description Test whether a worksheet exists in a workbook.
'@param sheetName String. Name to look up.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Boolean indicating existence.
Public Function WorksheetExists(ByVal sheetName As String, _
                                Optional ByVal targetBook As Workbook) As Boolean

    Dim wb As Workbook
    Dim sh As Worksheet

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not (sh Is Nothing)
End Function

'@description Remove data, tables, shapes and names from a worksheet.
'@param sh Worksheet to clear.
Public Sub ClearWorksheet(ByVal sh As Worksheet)

    Dim nm As Name

    If sh Is Nothing Then Exit Sub

    BusyApp

    On Error Resume Next
        Do While sh.ListObjects.Count > 0
            sh.ListObjects(1).Delete
        Loop

        Do While sh.Shapes.Count > 0
            sh.Shapes(1).Delete
        Loop

        For Each nm In sh.Names
            nm.Delete
        Next nm

        For Each nm In sh.Parent.Names
            If InStr(1, nm.RefersTo, "'" & sh.Name & "'!", vbTextCompare) > 0 Then nm.Delete
        Next nm

        sh.Cells.Clear
    On Error GoTo 0
End Sub

'@section Named Ranges
'===============================================================================

'@description Determine whether a workbook or worksheet name exists.
'@param nameText String. Name to inspect.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Boolean True when the name is present.
Public Function NamedRangeExists(ByVal nameText As String, _
                                 Optional ByVal targetBook As Workbook) As Boolean

    Dim wb As Workbook
    Dim nm As Name

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set nm = wb.Names(nameText)
    On Error GoTo 0

    If Not (nm Is Nothing) Then
        NamedRangeExists = True
        Exit Function
    End If

    Dim sheetName As String
    sheetName = ParseSheetName(nameText)

    If sheetName <> vbNullString And WorksheetExists(sheetName, wb) Then
        On Error Resume Next
            Set nm = wb.Worksheets(sheetName).Names(nameText)
        On Error GoTo 0
        NamedRangeExists = Not (nm Is Nothing)
    End If
End Function

'@description Extract worksheet name from qualified references like Sheet1!Name.
'@param qualifiedName String possibly containing '!'.
'@return Worksheet name or empty string.
Private Function ParseSheetName(ByVal qualifiedName As String) As String
    Dim bangPos As Long

    bangPos = InStr(qualifiedName, "!")
    If bangPos > 0 Then
        ParseSheetName = Replace(Left$(qualifiedName, bangPos - 1), "'", vbNullString)
    End If
End Function

'@section Range Writers
'===============================================================================

'@description Write a row of values to a target range.
'@param target Range. Starting cell.
'@param values ParamArray values to write.
Public Sub WriteRow(ByVal target As Range, ParamArray values() As Variant)
    Dim idx As Long

    For idx = LBound(values) To UBound(values)
        target.Offset(0, idx - LBound(values)).Value = values(idx)
    Next idx
End Sub

'@description Write a column of values to a target range.
'@param target Range. Starting cell.
'@param values ParamArray values to write.
Public Sub WriteColumn(ByVal target As Range, ParamArray values() As Variant)
    Dim idx As Long

    For idx = LBound(values) To UBound(values)
        target.Offset(idx - LBound(values), 0).Value = values(idx)
    Next idx
End Sub

'@description Convert an array of row arrays into a 2D matrix.
'@param rows Variant. Array of arrays to convert.
'@return Variant 2D matrix or Empty when invalid.
Public Function RowsToMatrix(ByVal rows As Variant) As Variant
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim colLower As Long
    Dim colUpper As Long
    Dim r As Long
    Dim c As Long
    Dim matrix() As Variant

    If Not IsArray(rows) Then Exit Function

    rowLower = LBound(rows)
    rowUpper = UBound(rows)
    colLower = LBound(rows(rowLower))
    colUpper = UBound(rows(rowLower))

    ReDim matrix(1 To rowUpper - rowLower + 1, 1 To colUpper - colLower + 1)

    For r = rowLower To rowUpper
        For c = colLower To colUpper
            matrix(r - rowLower + 1, c - colLower + 1) = rows(r)(c)
        Next c
    Next r

    RowsToMatrix = matrix
End Function

'@description Write a 2D matrix into the supplied range.
'@param target Range. Upper-left cell for the matrix.
'@param matrix Variant. 2D array of values.
Public Sub WriteMatrix(ByVal target As Range, ByVal matrix As Variant)
    If IsEmpty(matrix) Then Exit Sub

    target.Resize(UBound(matrix, 1) - LBound(matrix, 1) + 1, _
                  UBound(matrix, 2) - LBound(matrix, 2) + 1).Value = matrix
End Sub

'@section Data Builders
'===============================================================================

'@description Create a BetterArray with the supplied items.
'@param items ParamArray values to push.
'@return BetterArray containing the items.
Public Function BetterArrayFromList(ParamArray items() As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.LowerBound = 0

    For idx = LBound(items) To UBound(items)
        result.Push items(idx)
    Next idx

    Set BetterArrayFromList = result
End Function

'@description Build a BetterArray from a 1-D Variant array.
'@param values Variant array.
'@return BetterArray with copied values.
Public Function BetterArrayFromVariant(ByVal values As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    If Not IsArray(values) Then Exit Function

    Set result = New BetterArray
    result.LowerBound = 0

    For idx = LBound(values) To UBound(values)
        result.Push values(idx)
    Next idx

    Set BetterArrayFromVariant = result
End Function

'@section Assertions
'===============================================================================

'@description Fail the current test when unexpected errors surface.
'@param assertObj Rubberduck Assert object.
'@param routineName String. Name of the failing routine.
Public Sub FailUnexpectedError(ByVal assertObj As Object, ByVal routineName As String)
    assertObj.Fail "Unexpected error in " & routineName & ": " & Err.Number & " - " & Err.Description
End Sub
