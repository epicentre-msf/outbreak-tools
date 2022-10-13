VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Properties
Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = False) As Range: End Property
'Getters for the encapsulated data
Public Property Get StartRow() As Long: End Property
Public Property Get StartColumn() As Long: End Property
Public Property Get Wksh() As Worksheet: End Property

'Exposed methods
Public Function FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray: End Function
Public Function FiltersData(ByVal varData As BetterArray, _
                            ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray: End Function
Public Function ColumnExists(ByVal colName As String) As Boolean: End Function
Public Function ColumnIndex(ByVal colName As String, _
                            Optional ByVal inDataRange As Boolean = False, _
                            Optional shouldExist As Boolean = False) As Long
End Function
