VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisVariablesCacheStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ITableSpecsVariablesCache stub exposing predefined dictionary and variables instances")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITableSpecsVariablesCache

Private Type TVariablesCacheState
    linelistSpecs As ILinelistSpecs
    dictionary As ILLdictionary
    variables As ILLVariables
    invalidateCalled As Boolean
End Type

Private this As TVariablesCacheState

'Configure the dictionary and variables references returned by the cache.
Public Sub Configure(ByVal dictionary As ILLdictionary, ByVal variables As ILLVariables)
    Set this.dictionary = dictionary
    Set this.variables = variables
End Sub

'Retrieve whether Invalidate was invoked.
Public Property Get InvalidateCalled() As Boolean
    InvalidateCalled = this.invalidateCalled
End Property

'Assign the linelist specifications reference for tests.
Public Property Set LinelistSpecificationsStub(ByVal linelistSpecs As ILinelistSpecs)
    Set this.linelistSpecs = linelistSpecs
End Property

'Retrieve the linelist specifications reference for assertions.
Public Property Get LinelistSpecificationsStub() As ILinelistSpecs
    Set LinelistSpecificationsStub = this.linelistSpecs
End Property

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsVariablesCache_Create(ByVal linelistSpecs As ILinelistSpecs) As ITableSpecsVariablesCache
    Set ITableSpecsVariablesCache_Create = Me
End Function

Private Property Get ITableSpecsVariablesCache_Self() As ITableSpecsVariablesCache
    Set ITableSpecsVariablesCache_Self = Me
End Property

Private Property Get ITableSpecsVariablesCache_LinelistSpecifications() As ILinelistSpecs
    Set ITableSpecsVariablesCache_LinelistSpecifications = this.linelistSpecs
End Property

Private Property Set ITableSpecsVariablesCache_LinelistSpecifications(ByVal linelistSpecs As ILinelistSpecs)
    Set this.linelistSpecs = linelistSpecs
End Property

Private Property Get ITableSpecsVariablesCache_Dictionary() As ILLdictionary
    Set ITableSpecsVariablesCache_Dictionary = this.dictionary
End Property

Private Property Get ITableSpecsVariablesCache_Variables() As ILLVariables
    Set ITableSpecsVariablesCache_Variables = this.variables
End Property

Private Sub ITableSpecsVariablesCache_Invalidate()
    this.invalidateCalled = True
End Sub

Private Sub ITableSpecsVariablesCache_Refresh()
End Sub
