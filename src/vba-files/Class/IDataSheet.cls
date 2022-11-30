VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Procedures for data manipulation in the linelist"


'@Folder("Data")
'@ModuleDescription("Interface of DataSheet Object")
'@IgnoreModule

Option Explicit

'Data Range
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'Getters for the encapsulated data
Public Property Get StartRow() As Long: End Property

Public Property Get StartColumn() As Long: End Property
Public Property Get DataEndRow() As Long: End Property

Public Property Get DataEndColumn() As Long: End Property
Public Property Get Wksh() As Worksheet: End Property

'Exposed methods
'Check if a column exists
Public Function ColumnExists(ByVal colName As String) As Boolean: End Function

'Return the column index
Public Function ColumnIndex(ByVal colName As String, _
                            Optional ByVal inDataRange As Boolean = False, _
                            Optional shouldExist As Boolean = False) As Long
End Function

'simple filter
Public Function FilterData(ByVal varName As String, ByVal criteriaName As String, _
                           ByVal returnedColumnName As String, _
                           Optional ByVal includeHeaders As Boolean = False) As BetterArray
End Function

'multiple filters
Public Function FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
End Function

'Import and Export Sub
Public Sub Export(ByVal toWkb As Workbook, _
                  Optional ByVal filteredVarName As String = "__all__", _
                  Optional ByVal filteredCondition As String = "<>", Optional ByVal Hide As Long = xlSheetHidden)
End Sub

Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long): End Sub









