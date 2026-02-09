VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for cross-table building"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for cross-table building")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Properties
'===============================================================================

' @description Header range of the table on the output worksheet
Public Property Get HeaderRange() As Range: End Property

' @description Row categories range, with optional inclusion of headers and filtering
Public Property Get RowsCategoriesRange( _
    Optional ByVal includeHeaders As Boolean = True, _
    Optional ByVal onlyCategories As Boolean = False) As Range
End Property

' @description Column data range for a specific column value
Public Property Get ColumnRange(ByVal colVal As String, _
    Optional ByVal onlyCategories As Boolean = False, _
    Optional ByVal includeHeaders As Boolean = False) As Range
End Property

' @description Row data range for a specific row value
Public Property Get RowRange(ByVal rowVal As String, _
    Optional ByVal includeHeaders As Boolean = True) As Range
End Property

' @description Starting row of the table on the output worksheet
Public Property Get StartRow() As Long: End Property

' @description Previous cross-table in the same section
Public Property Get Previous() As ICrossTable: End Property

' @description Table specifications used to build this cross-table
Public Property Get Specifications() As ITableSpecs: End Property

' @description Output worksheet where the table is built
Public Property Get Wksh() As Worksheet: End Property

' @description Number of data columns in the table
Public Property Get NumberOfColumns() As Long: End Property

' @description Last column used by the table on the worksheet
Public Property Get EndColumn() As Long: End Property

' @description Last row used by the table on the worksheet
Public Property Get EndRow() As Long: End Property

' @description Header label for time series tables
Public Property Get TimeSeriesHeader() As String: End Property

' @description List of named range names created by this table
Public Property Get NamedRangesList() As BetterArray: End Property

'@section Methods
'===============================================================================

' @description Add row headers and row category labels to the worksheet
Public Sub AddRows(): End Sub

' @description Add column headers and column category labels to the worksheet
Public Sub AddColumns(): End Sub

' @description Add the table title/header to the worksheet
Public Sub AddHeader(): End Sub

' @description Create all named ranges for the table data regions
Public Sub NameRanges(): End Sub

' @description Build the complete table (AddHeader, AddRows, AddColumns, NameRanges)
Public Sub Build(): End Sub

' @description Apply formatting to the table using design specifications
Public Sub Format(ByVal designFormat As ILLFormat): End Sub
