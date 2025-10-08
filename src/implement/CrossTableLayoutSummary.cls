VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableLayoutSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Value object describing cross table layout coordinates"

'@Folder("Implement")
'@ModuleDescription("Immutable container describing start offsets and spans for cross tables")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableLayoutSummary
    startRow As Long
    startColumn As Long
    headerRowCount As Long
    dataRowCount As Long
    dataColumnCount As Long
End Type

Private this As TCrossTableLayoutSummary

'@section Public API
'===============================================================================

Public Sub Initialise(ByVal startRow As Long, _
                      ByVal startColumn As Long, _
                      ByVal headerRowCount As Long, _
                      ByVal dataRowCount As Long, _
                      ByVal dataColumnCount As Long)
    this.startRow = startRow
    this.startColumn = startColumn
    this.headerRowCount = headerRowCount
    this.dataRowCount = dataRowCount
    this.dataColumnCount = dataColumnCount
End Sub

Public Property Get StartRow() As Long
    StartRow = this.startRow
End Property

Public Property Get StartColumn() As Long
    StartColumn = this.startColumn
End Property

Public Property Get HeaderRowCount() As Long
    HeaderRowCount = this.headerRowCount
End Property

Public Property Get DataRowCount() As Long
    DataRowCount = this.dataRowCount
End Property

Public Property Get DataColumnCount() As Long
    DataColumnCount = this.dataColumnCount
End Property

End Class
