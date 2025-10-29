VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToEntryCollectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Utility builder accumulating AnalysisGoToEntry instances grouped by scope")
'@depends ProjectError, AnalysisGoToEntry
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TEntryCollectionState
    sections As BetterArray
    headers As BetterArray
    graphs As BetterArray
End Type

Private this As TEntryCollectionState

Private Const CLASSNAME As String = "AnalysisGoToEntryCollectionBuilder"

Private Sub Class_Initialize()
    Set this.sections = CreateBuffer()
    Set this.headers = CreateBuffer()
    Set this.graphs = CreateBuffer()
End Sub

'@section Helper factory
'===============================================================================
Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateBuffer = buffer
End Function

Private Sub AddEntryToTarget(ByVal target As BetterArray, ByVal entry As IAnalysisGoToEntry)
    If target Is Nothing Then Exit Sub
    target.Push entry
End Sub

Private Function CreateEntry(ByVal scope As String, ByVal labelText As String, ByVal prefixText As String, ByVal suffixText As String) As IAnalysisGoToEntry
    Set CreateEntry = AnalysisGoToEntry.Create(scope, labelText, prefixText, suffixText)
End Function

Private Sub ValidateEntry(ByVal entry As IAnalysisGoToEntry)
    If entry Is Nothing Then
        Err.Raise CLng(ProjectError.InvalidArgument), CLASSNAME, "entry cannot be Nothing"
    End If
End Sub

Private Sub RouteEntry(ByVal entry As IAnalysisGoToEntry)
    Dim scopeKey As String

    ValidateEntry entry

    scopeKey = LCase$(Trim$(entry.Scope))

    Select Case scopeKey
        Case "section"
            AddEntryToTarget this.sections, entry
        Case "header"
            AddEntryToTarget this.headers, entry
        Case "graph"
            AddEntryToTarget this.graphs, entry
        Case Else
            Err.Raise CLng(ProjectError.InvalidArgument), CLASSNAME, "Unsupported GoTo scope: " & entry.Scope
    End Select
End Sub

'@section Builders
'===============================================================================
Public Sub AddSection(ByVal labelText As String, Optional ByVal prefixText As String = vbNullString, Optional ByVal suffixText As String = vbNullString)
    RouteEntry CreateEntry("section", labelText, prefixText, suffixText)
End Sub

Public Sub AddHeader(ByVal labelText As String, Optional ByVal prefixText As String = vbNullString, Optional ByVal suffixText As String = vbNullString)
    RouteEntry CreateEntry("header", labelText, prefixText, suffixText)
End Sub

Public Sub AddGraph(ByVal labelText As String, Optional ByVal prefixText As String = vbNullString, Optional ByVal suffixText As String = vbNullString)
    RouteEntry CreateEntry("graph", labelText, prefixText, suffixText)
End Sub

Public Sub AddEntry(ByVal entry As IAnalysisGoToEntry)
    RouteEntry entry
End Sub

Public Sub AddEntries(ByVal entries As BetterArray)
    Dim idx As Long
    Dim entry As IAnalysisGoToEntry

    If entries Is Nothing Then Exit Sub
    If entries.Length = 0 Then Exit Sub

    For idx = entries.LowerBound To entries.UpperBound
        Set entry = entries.Item(idx)
        If Not entry Is Nothing Then RouteEntry entry
    Next idx
End Sub

'@section Accessors
'===============================================================================
Public Property Get SectionEntries() As BetterArray
    Set SectionEntries = this.sections
End Property

Public Property Get HeaderEntries() As BetterArray
    Set HeaderEntries = this.headers
End Property

Public Property Get GraphEntries() As BetterArray
    Set GraphEntries = this.graphs
End Property
