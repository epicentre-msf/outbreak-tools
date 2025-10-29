VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetGraphInputsAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Adapts worksheet inputs to the graph engine input interface")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphEngineInputs

Private Type TAdapterState
    worksheetInputs As IAnalysisWorksheetInputs
    graphSpec As IAnalysisWorksheetGraphSpec
End Type

Private this As TAdapterState

Private Const CLASSNAME As String = "AnalysisWorksheetGraphInputsAdapter"

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal worksheetInputs As IAnalysisWorksheetInputs, _
                      ByVal graphSpec As IAnalysisWorksheetGraphSpec)

    ValidateInputs worksheetInputs
    ValidateGraphSpec graphSpec

    Set this.worksheetInputs = worksheetInputs
    Set this.graphSpec = graphSpec
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateInputs(ByVal worksheetInputs As IAnalysisWorksheetInputs)
    If worksheetInputs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "worksheetInputs cannot be Nothing"
    End If
End Sub

Private Sub ValidateGraphSpec(ByVal graphSpec As IAnalysisWorksheetGraphSpec)
    If graphSpec Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "graphSpec cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisGraphEngineInputs_CrossTableFor(ByVal planItem As IAnalysisGraphPlanItem) As ICrossTable
    'Cross-table graphs are not handled by this adapter; defer to upstream logic when required.
    Set IAnalysisGraphEngineInputs_CrossTableFor = Nothing
End Function

Private Property Get IAnalysisGraphEngineInputs_GraphList() As ListObject
    Dim graphId As String

    graphId = this.graphSpec.GraphListId
    If LenB(graphId) = 0 Then
        Set IAnalysisGraphEngineInputs_GraphList = Nothing
    Else
        Set IAnalysisGraphEngineInputs_GraphList = this.worksheetInputs.SpecificationList(graphId)
    End If
End Property

Private Property Get IAnalysisGraphEngineInputs_TimeSeriesList() As ListObject
    Dim tableId As String

    tableId = this.graphSpec.TableId
    If LenB(tableId) = 0 Then
        Set IAnalysisGraphEngineInputs_TimeSeriesList = Nothing
    Else
        Set IAnalysisGraphEngineInputs_TimeSeriesList = this.worksheetInputs.SpecificationList(tableId)
    End If
End Property

Private Property Get IAnalysisGraphEngineInputs_TitleList() As ListObject
    Dim listId As String

    listId = this.graphSpec.GraphTitleListId
    If LenB(listId) = 0 Then
        Set IAnalysisGraphEngineInputs_TitleList = Nothing
    Else
        Set IAnalysisGraphEngineInputs_TitleList = this.worksheetInputs.SpecificationList(listId)
    End If
End Property

Private Property Get IAnalysisGraphEngineInputs_Linelist() As ILinelistSpecs
    Set IAnalysisGraphEngineInputs_Linelist = this.worksheetInputs.Linelist
End Property
