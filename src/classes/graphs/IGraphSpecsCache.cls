VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSpecsCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Caches graph list object rows keyed by graph identifier")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a cache bound to the provided graph list object.
Public Function Create(ByVal graphList As ListObject) As IGraphSpecsCache
End Function

'Return the current instance as its interface type.
Public Property Get Self() As IGraphSpecsCache
End Property

'@section Accessors
'===============================================================================
'Retrieve the underlying graph list object.
Public Property Get GraphList() As ListObject
End Property

'Assign a new graph list object and invalidate cached data.
Public Property Set GraphList(ByVal graphList As ListObject)
End Property

'Return the cached list of graph identifiers.
Public Function GraphIds() As BetterArray
End Function

'Return the values recorded for a column within a specific graph.
Public Function ColumnValues(ByVal graphId As String, ByVal columnName As String) As BetterArray
End Function

'@section Cache management
'===============================================================================
'Clear cached data so the next access rebuilds it.
Public Sub Invalidate()
End Sub

'Refresh cached data immediately.
Public Sub Refresh()
End Sub

