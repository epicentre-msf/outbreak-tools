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
Public Property Get Column(ByVal columnName As String) As BetterArray: End Property
Public Property Get ColumnExists(ByVal columnName As String) As Boolean: End Property
Public Property Get FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String) As BetterArray
End Property
Public Property Get FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
End Property

'Getters for the encapsulated data
Public Property Get StartLine() As Long: End Property
Public Property Get StartColumn() As Long: End Property
