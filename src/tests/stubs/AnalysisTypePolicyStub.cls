VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTypePolicyStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ITableTypePolicy test double returning configurable evaluation flags")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITableTypePolicy

Private Type TPolicyStubState
    tableTypeCode As Byte
    isValid As Boolean
    hasPercent As Boolean
    hasTotal As Boolean
    hasMissing As Boolean
    hasGraph As Boolean
End Type

Private this As TPolicyStubState

'Configure the stub behaviour for downstream invocations.
Public Sub Configure(ByVal tableTypeCode As Byte, _
                     ByVal isValid As Boolean, _
                     ByVal hasPercent As Boolean, _
                     ByVal hasTotal As Boolean, _
                     ByVal hasMissing As Boolean, _
                     ByVal hasGraph As Boolean)
    this.tableTypeCode = tableTypeCode
    this.isValid = isValid
    this.hasPercent = hasPercent
    this.hasTotal = hasTotal
    this.hasMissing = hasMissing
    this.hasGraph = hasGraph
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = this.tableTypeCode
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    '@Ignore context
    ITableTypePolicy_HasPercent = this.hasPercent
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    '@Ignore context
    ITableTypePolicy_HasTotal = this.hasTotal
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    '@Ignore context
    ITableTypePolicy_HasMissing = this.hasMissing
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    '@Ignore context
    ITableTypePolicy_HasGraph = this.hasGraph
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    '@Ignore context
    ITableTypePolicy_IsValid = this.isValid
End Function
