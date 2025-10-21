VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableSequenceBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableSequenceBuilder stub returning a preset BetterArray")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableSequenceBuilder

Private Type TSequenceStubState
    results As BetterArray
    lastSpecificationList As ListObject
    lastLinelistSpecs As ILinelistSpecs
    lastTableTypeCode As Byte
    buildCount As Long
End Type

Private this As TSequenceStubState

Private Sub Class_Initialize()
    Set this.results = New BetterArray
    this.results.LowerBound = 1
End Sub

'Configure the results returned by the stub.
Public Sub SetResults(ByVal results As BetterArray)
    Set this.results = results
End Sub

Public Property Get LastSpecificationList() As ListObject
    Set LastSpecificationList = this.lastSpecificationList
End Property

Public Property Get LastLinelistSpecs() As ILinelistSpecs
    Set LastLinelistSpecs = this.lastLinelistSpecs
End Property

Public Property Get LastTableTypeCode() As Byte
    LastTableTypeCode = this.lastTableTypeCode
End Property

Public Property Get BuildCount() As Long
    BuildCount = this.buildCount
End Property

Private Function IAnalysisTableSequenceBuilder_Build(ByVal specificationList As ListObject, _
                                                     ByVal linelistSpecs As ILinelistSpecs, _
                                                     ByVal tableTypeCode As Byte) As BetterArray
    Set this.lastSpecificationList = specificationList
    Set this.lastLinelistSpecs = linelistSpecs
    this.lastTableTypeCode = tableTypeCode
    this.buildCount = this.buildCount + 1
    Set IAnalysisTableSequenceBuilder_Build = this.results
End Function
