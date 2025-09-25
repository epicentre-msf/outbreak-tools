VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsColumnMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Contract for caching table specification column metadata")
'@depends BetterArray
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a new column map bound to the provided header and data ranges.
Public Function Create(ByVal headerRange As Range, _
                       ByVal tableRange As Range) As ITableSpecsColumnMap
End Function

'Retrieve the current instance as its interface type.
Public Property Get Self() As ITableSpecsColumnMap
End Property

'@section Ranges
'===============================================================================
'Read the header range used to prime the cache.
Public Property Get HeaderRange() As Range
End Property

'Update the header range that feeds the cache.
Public Property Set HeaderRange(ByVal headerRange As Range)
End Property

'Read the table/specification range that stores values.
Public Property Get TableRange() As Range
End Property

'Update the table/specification range that stores values.
Public Property Set TableRange(ByVal tableRange As Range)
End Property

'@section Lookups
'===============================================================================
'Fast existence check for a column name or fragment.
Public Function ColumnExists(ByVal columnName As String) As Boolean
End Function

'Get the one-based index of a column from the cached header.
Public Function ColumnIndex(ByVal columnName As String) As Long
End Function

'Return the value stored in the cached table range for the column.
Public Function Value(ByVal columnName As String) As String
End Function

'@section Cache management
'===============================================================================
'Mark the cache as dirty so the next access rebuilds metadata.
Public Sub Invalidate()
End Sub

'Rebuild the cache immediately using the current ranges.
Public Sub Refresh()
End Sub
