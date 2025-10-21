VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToDropdownRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable implementation of IAnalysisGoToDropdownRequest")
'@depends ProjectError, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGoToDropdownRequest

Private Type TDropdownRequestState
    scope As String
    listName As String
    cellName As String
    labelText As String
    entries As BetterArray
    values As BetterArray
End Type

Private this As TDropdownRequestState

Private Const CLASSNAME As String = "AnalysisGoToDropdownRequest"

'@section Factory
'===============================================================================
Public Function Create(ByVal scope As String, _
                       ByVal listName As String, _
                       ByVal cellName As String, _
                       ByVal labelText As String, _
                       ByVal entries As BetterArray, _
                       ByVal values As BetterArray) As IAnalysisGoToDropdownRequest

    ValidateRequiredText scope, "scope"
    ValidateRequiredText listName, "listName"
    ValidateRequiredText cellName, "cellName"
    ValidateRequiredText labelText, "labelText"
    ValidateArray entries, "entries"
    ValidateArray values, "values"

    With New AnalysisGoToDropdownRequest
        .Initialise scope, listName, cellName, labelText, entries, values
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IAnalysisGoToDropdownRequest
    Set Self = Me
End Property

Public Sub Initialise(ByVal scope As String, _
                      ByVal listName As String, _
                      ByVal cellName As String, _
                      ByVal labelText As String, _
                      ByVal entries As BetterArray, _
                      ByVal values As BetterArray)

    this.scope = Trim$(scope)
    this.listName = Trim$(listName)
    this.cellName = Trim$(cellName)
    this.labelText = labelText
    Set this.entries = AnalysisArrayUtilities.CloneItems(entries)
    Set this.values = AnalysisArrayUtilities.CloneItems(values)
End Sub

Private Sub ValidateRequiredText(ByVal valueText As String, ByVal argumentName As String)
    If LenB(Trim$(valueText)) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be empty"
    End If
End Sub

Private Sub ValidateArray(ByVal candidate As BetterArray, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Property Get HasInternalEntries() As Boolean
    If this.entries Is Nothing Then
        HasInternalEntries = False
    Else
        HasInternalEntries = (this.entries.Length > 0)
    End If
End Property

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGoToDropdownRequest_Scope() As String
    IAnalysisGoToDropdownRequest_Scope = this.scope
End Property

Private Property Get IAnalysisGoToDropdownRequest_ListName() As String
    IAnalysisGoToDropdownRequest_ListName = this.listName
End Property

Private Property Get IAnalysisGoToDropdownRequest_CellName() As String
    IAnalysisGoToDropdownRequest_CellName = this.cellName
End Property

Private Property Get IAnalysisGoToDropdownRequest_LabelText() As String
    IAnalysisGoToDropdownRequest_LabelText = this.labelText
End Property

Private Property Get IAnalysisGoToDropdownRequest_Entries() As BetterArray
    Set IAnalysisGoToDropdownRequest_Entries = AnalysisArrayUtilities.CloneItems(this.entries)
End Property

Private Property Get IAnalysisGoToDropdownRequest_Values() As BetterArray
    Set IAnalysisGoToDropdownRequest_Values = AnalysisArrayUtilities.CloneItems(this.values)
End Property

Private Property Get IAnalysisGoToDropdownRequest_HasEntries() As Boolean
    IAnalysisGoToDropdownRequest_HasEntries = HasInternalEntries()
End Property
