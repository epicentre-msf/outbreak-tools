VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToDropdownFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Creates dropdown requests from AnalysisGoToEntry collections")
'@depends ProjectError, AnalysisGoToDropdownRequest, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Const LIST_SUFFIX As String = "gotosection"
Private Const CELL_SUFFIX As String = "go_to_section"

'@section Factory helpers
'===============================================================================
Public Function CreateRequest(ByVal scope As String, _
                              ByVal entries As BetterArray, _
                              ByVal goToPrefix As String, _
                              ByVal goToSuffix As String, _
                              ByVal labelText As String) As IAnalysisGoToDropdownRequest

    ValidateEntries entries, scope

    Dim values As BetterArray
    Set values = BuildValues(entries)

    If values.Length = 0 Then Exit Function

    Dim listName As String
    Dim cellName As String

    listName = BuildListName(goToPrefix, goToSuffix)
    cellName = BuildCellName(goToPrefix, goToSuffix)

    Set CreateRequest = AnalysisGoToDropdownRequest.Create(scope, listName, cellName, labelText, entries, values)
End Function

Public Function CreateAll(ByVal builder As AnalysisGoToEntryCollectionBuilder, _
                          ByVal sectionPrefix As String, _
                          ByVal headerPrefix As String, _
                          ByVal graphPrefix As String, _
                          ByVal goToSuffix As String, _
                          ByVal sectionLabel As String, _
                          ByVal headerLabel As String, _
                          ByVal graphLabel As String) As BetterArray

    Dim requests As BetterArray
    Dim request As IAnalysisGoToDropdownRequest

    Set requests = AnalysisArrayUtilities.CreateBuffer()

    Set request = CreateRequest("section", builder.SectionEntries, sectionPrefix, goToSuffix, sectionLabel)
    If Not request Is Nothing Then requests.Push request

    Set request = CreateRequest("header", builder.HeaderEntries, headerPrefix, goToSuffix, headerLabel)
    If Not request Is Nothing Then requests.Push request

    Set request = CreateRequest("graph", builder.GraphEntries, graphPrefix, goToSuffix, graphLabel)
    If Not request Is Nothing Then requests.Push request

    Set CreateAll = requests
End Function

Private Function BuildValues(ByVal entries As BetterArray) As BetterArray
    Dim values As BetterArray
    Dim idx As Long
    Dim entry As IAnalysisGoToEntry

    Set values = AnalysisArrayUtilities.CreateBuffer()

    If entries Is Nothing Then
        Set BuildValues = values
        Exit Function
    End If

    If entries.Length = 0 Then
        Set BuildValues = values
        Exit Function
    End If

    For idx = entries.LowerBound To entries.UpperBound
        Set entry = entries.Item(idx)
        If Not entry Is Nothing Then values.Push entry.DisplayText
    Next idx

    Set BuildValues = values
End Function

Private Function BuildListName(ByVal prefix As String, ByVal suffix As String) As String
    BuildListName = Trim$(prefix) & LIST_SUFFIX & Trim$(suffix)
End Function

Private Function BuildCellName(ByVal prefix As String, ByVal suffix As String) As String
    BuildCellName = Trim$(prefix) & CELL_SUFFIX & Trim$(suffix)
End Function

Private Sub ValidateEntries(ByVal entries As BetterArray, ByVal scope As String)
    If entries Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "entries for scope " & scope & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, TypeName(Me), messageText
End Sub
