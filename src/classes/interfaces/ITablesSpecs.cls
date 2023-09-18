VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for table specifications")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Enum AnalysisTablesType
    TypeGlobalSummary = 1
    TypeUnivariate = 2
    TypeBivariate = 3
    TypeTimeSeries = 4
    TypeSpatial = 5
    TypeSpatioTemporal = 6
End Enum

'One of the previous mentioned table types
Public Property Get TableType() As AnalysisTablesType: End Property 
Public Property Get Dictionary() As ILLdictionary: End Property
Public Property Get RowCategories() As BetterArray: End Property
Public Property Get ColumnCategories() As BetterArray: End Property
Public Property Get isNewSection() As Boolean: End Property
Public Property Get TableId() As String: End Property
Public Property Get ValidTable() As Boolean: End Property
Public Property Get HasMissing() As Boolean: End Property
Public Property Get HasTotal() As Boolean: End Property
Public Property Get HasGraph() As Boolean: End Property
Public Property Get HasPercentage() As Boolean: End Property
'Find the type of spatial analysis
Public Property Get SpatialTableTypes() As String: End Property
'Return the previous table specification in the same section 
'(the order is defined by apearance in the listobject)
Public Property Get Previous() As ITablesSpecs: End Property
'Return the next valid table (anchorRng is the last range of the listObject)
Public Property Get NextSpecs(ByVal anchorRng As Range) As ITablesSpecs
End Property
'The id of the first table of the section
Public Property Get TableSectionId() As String: End Property 
'Find the value of a colname: The search is done without case sensitive
Public Function Value(ByVal colName As String) As String: End Function 

