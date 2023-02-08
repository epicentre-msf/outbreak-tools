VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("Interface for table specifications")
'@IgnoreModule

Option Explicit

Public Enum AnalysisTablesType
    TypeGlobalSummary = 1
    TypeUnivariate = 2
    TypeBivariate = 3
    TypeTimeSeries = 4
    TypeSpatial = 5
End Enum

Public Property Get TableType() As AnalysisTablesType: End Property 'One of the previous mentioned table types

Public Property Get Dictionary() As ILLdictionary: End Property
Public Property Get RowCategories() As BetterArray: End Property

Public Property Get ColumnCategories() As BetterArray: End Property
Public Property Get isNewSection() As Boolean: End Property

Public Property Get TableId() As String: End Property
Public Property Get ValidTable() As Boolean: End Property

Public Property Get HasMissing() As Boolean: End Property

Public Property Get HasTotal() As Boolean: End Property

Public Property Get HasPercentage() As Boolean: End Property

Public Property Get Previous() As ITablesSpecs: End Property 'Return the previous table specification in the same section (the order is defined by apearance in the listobject)

Public Property Get TableSectionId() As String: End Property 'The id of the first table of the section

'Exposed methods
Public Function Value(ByVal colName As String) As String: End Function 'The search is done without case sensitive

'Find the type of spatial analysis
Public Property Get SpatialTableTypes() As String
End Property

