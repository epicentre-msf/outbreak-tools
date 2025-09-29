Attribute VB_Name = "LLFormatTestFixture"
Attribute VB_Description = "Shared helpers for LLFormat tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Shared helpers for LLFormat tests using workbook fixtures")

Public Const LLFORMAT_TEMPLATE_SHEET As String = "LLFormatFixture"

Public Function PrepareLLFormatFixture(ByVal sheetName As String, _
                                       Optional ByVal targetBook As Workbook) As Worksheet

    Dim wb As Workbook
    Dim template As Worksheet
    Dim copySheet As Worksheet

    On Error GoTo Fail

    Set wb = ResolveWorkbook(targetBook)
    Set template = LLFormatTemplate(wb)

    DeleteLLFormatFixture sheetName, wb

    template.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Set copySheet = wb.Worksheets(wb.Worksheets.Count)
    copySheet.Name = sheetName

    Set PrepareLLFormatFixture = copySheet
    Exit Function

Fail:
    Err.Raise Err.Number, "LLFormatTestFixture.PrepareLLFormatFixture", Err.Description
End Function

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

Public Function LLFormatTemplate(Optional ByVal targetBook As Workbook) As Worksheet

    Dim wb As Workbook
    Dim template As Worksheet

    Set wb = ResolveWorkbook(targetBook)

    On Error Resume Next
        Set template = wb.Worksheets(LLFORMAT_TEMPLATE_SHEET)
    On Error GoTo 0

    If template Is Nothing Then
        Err.Raise vbObjectError + 513, "LLFormatTestFixture.LLFormatTemplate", _
                  "Worksheet '" & LLFORMAT_TEMPLATE_SHEET & "' is required for LLFormat tests"
    End If

    Set LLFormatTemplate = template
End Function

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

    If hostSheet.ListObjects.Count = 0 Then
        Err.Raise vbObjectError + 515, "LLFormatTestFixture.FixtureCell", _
                  "Fixture sheet '" & hostSheet.Name & "' must expose a format table"
    End If

    Set tableObj = hostSheet.ListObjects(1)
    Set labelRange = tableObj.ListColumns("label").DataBodyRange

    On Error Resume Next
        Set labelCell = labelRange.Find(What:=labelText, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If labelCell Is Nothing Then
        Err.Raise vbObjectError + 516, "LLFormatTestFixture.FixtureCell", _
                  "Label '" & labelText & "' is missing from fixture sheet '" & hostSheet.Name & "'"
    End If

    On Error Resume Next
        Set designRange = tableObj.ListColumns(designColumn).Range
    On Error GoTo 0

    If designRange Is Nothing Then
        Err.Raise vbObjectError + 517, "LLFormatTestFixture.FixtureCell", _
                  "Design column '" & designColumn & "' is missing from fixture sheet '" & hostSheet.Name & "'"
    End If

    columnIndex = designRange.Column
    Set FixtureCell = hostSheet.Cells(labelCell.Row, columnIndex)
End Function

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
