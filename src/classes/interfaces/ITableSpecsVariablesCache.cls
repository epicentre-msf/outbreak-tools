VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsVariablesCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Caches dictionary-driven variable lookups for TableSpecs helpers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a new cache bound to the supplied linelist specifications source.
Public Function Create(ByVal linelistSpecs As ILinelistSpecs) As ITableSpecsVariablesCache
End Function

'Retrieve the current instance cast as the interface type.
Public Property Get Self() As ITableSpecsVariablesCache
End Property

'@section Source configuration
'===============================================================================
'Retrieve the underlying linelist specifications reference.
Public Property Get LinelistSpecifications() As ILinelistSpecs
End Property

'Assign a new linelist specifications reference and reset cached collaborators.
Public Property Set LinelistSpecifications(ByVal linelistSpecs As ILinelistSpecs)
End Property

'@section Cached accessors
'===============================================================================
'Return the dictionary associated with the bound linelist specifications.
Public Property Get Dictionary() As ILLdictionary
End Property

'Return a cached instance of ILLVariables tied to the dictionary.
Public Property Get Variables() As ILLVariables
End Property

'@section Cache management
'===============================================================================
'Clear cached references so the next lookup rebuilds them.
Public Sub Invalidate()
End Sub

'Force refresh of cached collaborators immediately.
Public Sub Refresh()
End Sub
