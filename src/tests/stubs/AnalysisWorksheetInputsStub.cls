VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetInputsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IAnalysisWorksheetInputs for coordinator tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisWorksheetInputs

Private Type TInputsStubState
    linelist As ILinelistSpecs
    outputSheets As Collection
    specificationLists As Collection
    requestedListIds As Collection
End Type

Private this As TInputsStubState

Private Sub Class_Initialize()
    Set this.outputSheets = New Collection
    Set this.specificationLists = New Collection
    Set this.requestedListIds = New Collection
End Sub

Public Sub ConfigureLinelist(ByVal linelist As ILinelistSpecs)
    Set this.linelist = linelist
End Sub

Public Sub RegisterOutputSheet(ByVal scope As AnalysisScope, ByVal sheet As Worksheet)
    On Error Resume Next
        this.outputSheets.Remove CStr(scope)
    On Error GoTo 0

    this.outputSheets.Add sheet, CStr(scope)
End Sub

Public Sub RegisterSpecificationList(ByVal listId As String, ByVal listObject As ListObject)
    Dim key As String
    key = LCase$(Trim$(listId))

    On Error Resume Next
        this.specificationLists.Remove key
    On Error GoTo 0

    this.specificationLists.Add listObject, key
End Sub

Public Property Get RequestedListIds() As Collection
    Set RequestedListIds = this.requestedListIds
End Property

Private Function IAnalysisWorksheetInputs_Linelist() As ILinelistSpecs
    Set IAnalysisWorksheetInputs_Linelist = this.linelist
End Function

Private Function IAnalysisWorksheetInputs_OutputSheet(ByVal scope As AnalysisScope) As Worksheet
    On Error Resume Next
        Set IAnalysisWorksheetInputs_OutputSheet = this.outputSheets.Item(CStr(scope))
    On Error GoTo 0
End Function

Private Function IAnalysisWorksheetInputs_SpecificationList(ByVal listId As String) As ListObject
    Dim key As String

    key = LCase$(Trim$(listId))
    this.requestedListIds.Add key

    On Error Resume Next
        Set IAnalysisWorksheetInputs_SpecificationList = this.specificationLists.Item(key)
    On Error GoTo 0
End Function

Private Property Get IAnalysisWorksheetInputs_Translator() As ITranslationObject
    Set IAnalysisWorksheetInputs_Translator = Nothing
End Property

Private Property Get IAnalysisWorksheetInputs_NamesRegistry() As IAnaTabIds
    Set IAnalysisWorksheetInputs_NamesRegistry = Nothing
End Property

Private Property Get IAnalysisWorksheetInputs_DesignFormatter() As ILLFormat
    Set IAnalysisWorksheetInputs_DesignFormatter = Nothing
End Property

Private Function IAnalysisWorksheetInputs_GoToLabel(ByVal scope As String) As String
    IAnalysisWorksheetInputs_GoToLabel = scope & ": "
End Function
