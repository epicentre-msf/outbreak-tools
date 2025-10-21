VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableEnumeratorStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableEnumerator double returning a predefined sequence of iteration items")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableEnumerator

Private Type TEnumeratorStubState
    items As BetterArray
    lastSpecificationList As ListObject
    enumerateCount As Long
End Type

Private this As TEnumeratorStubState

Private Sub Class_Initialize()
    Set this.items = New BetterArray
    this.items.LowerBound = 1
End Sub

'Inject the iteration items that should be returned to consumers.
Public Sub SetItems(ByVal items As BetterArray)
    Set this.items = items
End Sub

'Expose the specification list received during enumeration for assertions.
Public Property Get LastSpecificationList() As ListObject
    Set LastSpecificationList = this.lastSpecificationList
End Property

'Expose how many times enumeration has been requested.
Public Property Get EnumerateCount() As Long
    EnumerateCount = this.enumerateCount
End Property

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTableEnumerator_Enumerate(ByVal specificationList As ListObject, _
                                                    ByVal linelistSpecs As ILinelistSpecs, _
                                                    ByVal tableTypeCode As Byte) As BetterArray
    '@Ignore linelistSpecs, tableTypeCode
    Set this.lastSpecificationList = specificationList
    this.enumerateCount = this.enumerateCount + 1
    Set IAnalysisTableEnumerator_Enumerate = this.items
End Function
