VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableEnumeratorBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ITableSpecsBuilder double feeding predictable specs to AnalysisTableEnumerator tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITableSpecsBuilder

Private Type TBuilderState
    rowValidity As Collection
    tableType As AnalysisTablesType
End Type

Private this As TBuilderState

Private Sub Class_Initialize()
    Set this.rowValidity = New Collection
    this.tableType = TypeUnivariate
End Sub

'@section Configuration
'===============================================================================
'Assign the table type returned by newly created stubs.
Public Sub SetTableType(ByVal tableType As AnalysisTablesType)
    this.tableType = tableType
End Sub

'Configure whether a given worksheet row should yield a valid TableSpecs instance.
Public Sub SetRowValidity(ByVal rowNumber As Long, ByVal isValid As Boolean)
    Dim key As String

    key = CStr(rowNumber)

    On Error Resume Next
        this.rowValidity.Remove key
    On Error GoTo 0

    On Error Resume Next
        this.rowValidity.Add isValid, key
    On Error GoTo 0
End Sub

'Expose the stub as its interface for convenient consumption.
Public Property Get Self() As ITableSpecsBuilder
    Set Self = Me
End Property

'@section Helpers
'===============================================================================
Private Function ResolveValidity(ByVal rowNumber As Long) As Boolean
    Dim key As String
    Dim stored As Variant

    key = CStr(rowNumber)
    ResolveValidity = True

    On Error Resume Next
        stored = this.rowValidity.Item(key)
    On Error GoTo 0

    If Not IsEmpty(stored) Then
        ResolveValidity = CBool(stored)
    End If
End Function

Private Function CreateSpecStub(ByVal tableRange As Range) As ITablesSpecs
    Dim stub As TableSpecsNavigatorSpecStub
    Dim validFlag As Boolean

    validFlag = ResolveValidity(tableRange.Row)
    Set stub = New TableSpecsNavigatorSpecStub
    stub.Configure validFlag, tableRange, this.tableType
    Set CreateSpecStub = stub.Self
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsBuilder_CreateSpec(ByVal headerRange As Range, _
                                               ByVal tableRange As Range, _
                                               ByVal linelistSpecs As ILinelistSpecs) As ITablesSpecs
    '@Ignore headerRange, linelistSpecs
    Set ITableSpecsBuilder_CreateSpec = CreateSpecStub(tableRange)
End Function
