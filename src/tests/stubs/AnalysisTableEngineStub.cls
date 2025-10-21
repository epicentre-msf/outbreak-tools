VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableEngineStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableEngine stub capturing invocations for coordinator tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableEngine

Private Type TEngineStubState
    runCount As Long
    tableTypeCodes As BetterArray
    sectionPrefixes As BetterArray
    headerPrefixes As BetterArray
End Type

Private this As TEngineStubState

Private Sub Class_Initialize()
    Set this.tableTypeCodes = CreateBuffer()
    Set this.sectionPrefixes = CreateBuffer()
    Set this.headerPrefixes = CreateBuffer()
End Sub

Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateBuffer = buffer
End Function

Public Property Get RunCount() As Long
    RunCount = this.runCount
End Property

Public Property Get TableTypeCodes() As BetterArray
    Set TableTypeCodes = this.tableTypeCodes
End Property

Public Property Get SectionPrefixes() As BetterArray
    Set SectionPrefixes = this.sectionPrefixes
End Property

Public Property Get HeaderPrefixes() As BetterArray
    Set HeaderPrefixes = this.headerPrefixes
End Property

Private Sub IAnalysisTableEngine_Run(ByVal specificationList As ListObject, _
                                     ByVal outputSheet As Worksheet, _
                                     ByVal linelistSpecs As ILinelistSpecs, _
                                     ByVal tableTypeCode As Byte, _
                                     ByVal writerContext As IAnalysisTableWriterContext, _
                                     Optional ByVal sectionLabelPrefix As String = vbNullString, _
                                     Optional ByVal headerLabelPrefix As String = vbNullString)

    Dim policyStub As AnalysisTablePolicyResultStub
    Dim iterationStub As AnalysisTableIterationItemStub
    Dim specStub As GraphTablesSpecsStub
    Dim planItem As IAnalysisTablePlanItem
    Dim sectionEntries As BetterArray
    Dim headerEntries As BetterArray

    this.runCount = this.runCount + 1
    this.tableTypeCodes.Push tableTypeCode
    this.sectionPrefixes.Push sectionLabelPrefix
    this.headerPrefixes.Push headerLabelPrefix

    If writerContext Is Nothing Then Exit Sub

    writerContext.BeginPlan outputSheet

    Set policyStub = New AnalysisTablePolicyResultStub
    Set iterationStub = New AnalysisTableIterationItemStub
    Set specStub = New GraphTablesSpecsStub
    specStub.Configure TypeUnivariate, "table_stub"
    iterationStub.Configure specStub.Self, False
    policyStub.Configure iterationStub, True
    policyStub.SetFlags False, False, False, True

    Set planItem = AnalysisTablePlanItem.Create(policyStub, 0)
    writerContext.WriteTable planItem, outputSheet

    Set sectionEntries = New BetterArray
    sectionEntries.LowerBound = 1
    sectionEntries.Push AnalysisGoToEntry.Create("section", "Section 1", sectionLabelPrefix)

    Set headerEntries = New BetterArray
    headerEntries.LowerBound = 1
    headerEntries.Push AnalysisGoToEntry.Create("header", "Header 1", headerLabelPrefix)

    writerContext.ApplyNavigation sectionEntries, headerEntries, outputSheet
    writerContext.CompletePlan outputSheet
End Sub
