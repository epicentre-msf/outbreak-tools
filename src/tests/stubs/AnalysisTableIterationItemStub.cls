VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableIterationItemStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableIterationItem stub providing minimal spec metadata")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableIterationItem

Private Type TIterationStubState
    specification As ITablesSpecs
    isNewSection As Boolean
End Type

Private this As TIterationStubState

Public Sub Configure(ByVal specification As ITablesSpecs, ByVal isNewSection As Boolean)
    Set this.specification = specification
    this.isNewSection = isNewSection
End Sub

Private Property Get IAnalysisTableIterationItem_Specification() As ITablesSpecs
    Set IAnalysisTableIterationItem_Specification = this.specification
End Property

Private Property Get IAnalysisTableIterationItem_IsNewSection() As Boolean
    IAnalysisTableIterationItem_IsNewSection = this.isNewSection
End Property

Private Property Get IAnalysisTableIterationItem_HeaderRange() As Range
    Set IAnalysisTableIterationItem_HeaderRange = Nothing
End Property

Private Property Get IAnalysisTableIterationItem_RowRange() As Range
    Set IAnalysisTableIterationItem_RowRange = Nothing
End Property

Private Property Get IAnalysisTableIterationItem_ListRow() As ListRow
    Set IAnalysisTableIterationItem_ListRow = Nothing
End Property

Private Property Get IAnalysisTableIterationItem_PreviousSpecification() As ITablesSpecs
    Set IAnalysisTableIterationItem_PreviousSpecification = Nothing
End Property

Private Property Get IAnalysisTableIterationItem_NextSpecification() As ITablesSpecs
    Set IAnalysisTableIterationItem_NextSpecification = Nothing
End Property
