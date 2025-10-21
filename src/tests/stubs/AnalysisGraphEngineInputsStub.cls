VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphEngineInputsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IAnalysisGraphEngineInputs for unit testing")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisGraphEngineInputs

Private Type TInputsStubState
    crossTables As Collection
    graphList As ListObject
    timeSeriesList As ListObject
    titleList As ListObject
    linelist As ILinelistSpecs
End Type

Private this As TInputsStubState

Private Sub Class_Initialize()
    Set this.crossTables = New Collection
End Sub

Public Sub ConfigureCrossTable(ByVal tableId As String, ByVal crossTable As ICrossTable)
    Dim key As String

    key = NormalizeKey(tableId)

    On Error Resume Next
        this.crossTables.Remove key
    On Error GoTo 0

    this.crossTables.Add crossTable, key
End Sub

Private Function NormalizeKey(ByVal tableId As String) As String
    NormalizeKey = Trim$(LCase$(tableId))
End Function

Public Property Let GraphList(ByVal graphList As ListObject)
    Set this.graphList = graphList
End Property

Public Property Get GraphList() As ListObject
    Set GraphList = this.graphList
End Property

Public Property Let TimeSeriesList(ByVal timeSeriesList As ListObject)
    Set this.timeSeriesList = timeSeriesList
End Property

Public Property Get TimeSeriesList() As ListObject
    Set TimeSeriesList = this.timeSeriesList
End Property

Public Property Let TitleList(ByVal titleList As ListObject)
    Set this.titleList = titleList
End Property

Public Property Get TitleList() As ListObject
    Set TitleList = this.titleList
End Property

Public Property Set Linelist(ByVal linelist As ILinelistSpecs)
    Set this.linelist = linelist
End Property

Public Property Get Linelist() As ILinelistSpecs
    Set Linelist = this.linelist
End Property

'@section Interface implementation
'===============================================================================
Private Function IAnalysisGraphEngineInputs_CrossTableFor(ByVal planItem As IAnalysisGraphPlanItem) As ICrossTable
    Dim tableId As String
    Dim key As String
    Dim crossTable As ICrossTable

    On Error Resume Next
        tableId = planItem.Specification.TableId
    On Error GoTo 0

    key = NormalizeKey(tableId)

    On Error Resume Next
        Set crossTable = this.crossTables.Item(key)
    On Error GoTo 0

    Set IAnalysisGraphEngineInputs_CrossTableFor = crossTable
End Function

Private Property Get IAnalysisGraphEngineInputs_GraphList() As ListObject
    Set IAnalysisGraphEngineInputs_GraphList = this.graphList
End Property

Private Property Get IAnalysisGraphEngineInputs_TimeSeriesList() As ListObject
    Set IAnalysisGraphEngineInputs_TimeSeriesList = this.timeSeriesList
End Property

Private Property Get IAnalysisGraphEngineInputs_TitleList() As ListObject
    Set IAnalysisGraphEngineInputs_TitleList = this.titleList
End Property

Private Property Get IAnalysisGraphEngineInputs_Linelist() As ILinelistSpecs
    Set IAnalysisGraphEngineInputs_Linelist = this.linelist
End Property
