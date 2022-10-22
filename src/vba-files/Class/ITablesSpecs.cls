VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ITablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Analysis")
'@ModuleDescription("Interface for talbe specifications")
'@IgnoreModule

Option Explicit

Public Enum AnalysisTablesType
  TypeGlobalSummary = 1
  TypeUnivariate = 2
  TypeBivariate = 3
  TypeTimeSeries = 4
  TypeSpatial = 4
End Enum

Public Property Get TableType() As AnalysisTablesType : End Property
Public Property Get RowCategories() As BetterArray : End Property
Public Property Get ColumnCategories() As BetterArray : End Property
Public Property Get isNewSection() As Boolean: End Property


'Exposed methods
Public Function Value(ByVal colName As String) As String : End Function

