VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Dictionary Class"
'@Folder("Dictionary")
'@ModuleDescription("Dictionary class for the analysis")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Column(ByVal columnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray: End Property
Public Property Get ColumnExists(ByVal columnName As String) As Boolean: End Property
Public Property Get FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
End Property
Public Property Get FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
End Property
Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = False) As Range: End Property
Public Property Get Prepared() As Boolean: End Property
Public Property Get Wksh() As Worksheet: End Property

'Getters for the encapsulated data
Public Property Get StartRow() As Long: End Property
Public Property Get StartColumn() As Long: End Property

'Void methods
Public Sub Prepare(Optional ByVal tablePrefix As String = "table"): End Sub
Public Sub Clean(): End Sub
