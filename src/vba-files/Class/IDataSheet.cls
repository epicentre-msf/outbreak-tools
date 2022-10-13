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
Public Function ColumnExists(ByVal columnName As String) As Boolean: End Function

'simple filter
Public Function FilterData(ByVal varName As String, ByVal criteriaName As String, _
                           ByVal returnedColumnName As String, _
                           Optional ByVal includeHeaders As Boolean = False) As BetterArray
End Function

'multiple filters
Public Function FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
End Function
