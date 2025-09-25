VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeriesSpecsFactoryStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stubbed ITimeSeriesTableSpecsFactory returning pre-registered specs per series id")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITimeSeriesTableSpecsFactory

Private Type TFactoryStub
    tableMap As Collection
End Type

Private this As TFactoryStub

Public Sub AddSpec(ByVal seriesId As String, ByVal specs As ITablesSpecs)
    Dim key As String

    key = NormalizeGraphKey(seriesId)
    If key = vbNullString Then Exit Sub

    EnsureMap

    On Error Resume Next
        this.tableMap.Add specs, key
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EnsureMap()
    If this.tableMap Is Nothing Then Set this.tableMap = New Collection
End Sub

Private Function Lookup(ByVal seriesId As String) As ITablesSpecs
    Dim key As String

    If this.tableMap Is Nothing Then Exit Function

    key = NormalizeGraphKey(seriesId)
    If key = vbNullString Then Exit Function

    On Error Resume Next
        Set Lookup = this.tableMap.Item(key)
    On Error GoTo 0
End Function

Private Function ITimeSeriesTableSpecsFactory_Create(ByVal headerRange As Range, _
                                                     ByVal listRow As ListRow, _
                                                     ByVal linelist As ILinelistSpecs) As ITablesSpecs
    Dim seriesId As String

    If Not listRow Is Nothing Then seriesId = CStr(listRow.Range.Cells(1, 1).Value)
    Set ITimeSeriesTableSpecsFactory_Create = Lookup(seriesId)
End Function

Private Property Get ITimeSeriesTableSpecsFactory_Self() As ITimeSeriesTableSpecsFactory
    Set ITimeSeriesTableSpecsFactory_Self = Me
End Property

