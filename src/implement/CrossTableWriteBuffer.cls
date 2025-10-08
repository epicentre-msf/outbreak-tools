VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableWriteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Batches cross table header/body output before writing to worksheets"

'@Folder("Implement")
'@ModuleDescription("Buffer for cross table header/body values prior to worksheet commit")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableWriteBuffer
    headerMatrix As Variant
    bodyMatrix As Variant
    startRow As Long
    startColumn As Long
End Type

Private this As TCrossTableWriteBuffer
Private Const CLASSNAME As String = "CrossTableWriteBuffer"

Public Sub Initialise(ByVal startRow As Long, ByVal startColumn As Long)
    this.startRow = startRow
    this.startColumn = startColumn
End Sub

Public Sub AssignHeader(values As Variant)
    this.headerMatrix = values
End Sub

Public Sub AssignBody(values As Variant)
    this.bodyMatrix = values
End Sub

Public Sub Commit(ByVal targetSheet As Worksheet)
    Dim nextRow As Long

    If targetSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Target worksheet is required"
    End If

    If this.startRow = 0 Or this.startColumn = 0 Then
        Err.Raise ProjectError.ErrorUnexpectedState, CLASSNAME, "Write buffer must be initialised before commit"
    End If

    nextRow = this.startRow

    'Write header rows first so body rows append immediately afterwards.
    If IsArray(this.headerMatrix) Then
        Dim headerRows As Long
        Dim headerCols As Long
        headerRows = MatrixRowCount(this.headerMatrix)
        headerCols = MatrixColumnCount(this.headerMatrix)

        If headerRows > 0 And headerCols > 0 Then
            targetSheet.Cells(nextRow, this.startColumn).Resize(headerRows, headerCols).Value = this.headerMatrix
            nextRow = nextRow + headerRows
        End If
    End If

    If IsArray(this.bodyMatrix) Then
        Dim bodyRows As Long
        Dim bodyCols As Long
        bodyRows = MatrixRowCount(this.bodyMatrix)
        bodyCols = MatrixColumnCount(this.bodyMatrix)

        If bodyRows > 0 And bodyCols > 0 Then
            targetSheet.Cells(nextRow, this.startColumn).Resize(bodyRows, bodyCols).Value = this.bodyMatrix
        End If
    End If
End Sub

Private Function MatrixRowCount(matrix As Variant) As Long
    On Error Resume Next
        MatrixRowCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
    On Error GoTo 0
End Function

Private Function MatrixColumnCount(matrix As Variant) As Long
    On Error Resume Next
        MatrixColumnCount = UBound(matrix, 2) - LBound(matrix, 2) + 1
    On Error GoTo 0
End Function

End Class
