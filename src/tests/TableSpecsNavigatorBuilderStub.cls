VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsNavigatorBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub builder used to test TableSpecsNavigator")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITableSpecsBuilder

Private Type TBuilderState
    rowValidity As Collection
    defaultTableType As AnalysisTablesType
End Type

Private this As TBuilderState

Private Sub Class_Initialize()
    Set this.rowValidity = New Collection
    this.defaultTableType = TypeUnivariate
End Sub

'Configure the default table type returned by created specs.
Public Sub SetTableType(ByVal tableType As AnalysisTablesType)
    this.defaultTableType = tableType
End Sub

'Mark a specific row as valid or invalid.
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

Private Function RowIsValid(ByVal rowNumber As Long) As Boolean
    Dim key As String
    Dim stored As Variant

    key = CStr(rowNumber)
    RowIsValid = False

    On Error Resume Next
        stored = this.rowValidity.Item(key)
    On Error GoTo 0

    If Not IsEmpty(stored) Then
        RowIsValid = CBool(stored)
    End If
End Function

Private Function ITableSpecsBuilder_CreateSpec(ByVal headerRange As Range, _
                                               ByVal tableRange As Range, _
                                               ByVal linelistSpecs As ILinelistSpecs) As ITablesSpecs
    Dim stub As TableSpecsNavigatorSpecStub

    Set stub = New TableSpecsNavigatorSpecStub
    stub.Configure RowIsValid(tableRange.Row), tableRange, this.defaultTableType
    Set ITableSpecsBuilder_CreateSpec = stub.Self
End Function

