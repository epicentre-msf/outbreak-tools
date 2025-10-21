VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToDropdownApplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Applies dropdown requests to IDropdownLists and target cells")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Sub Apply(ByVal request As IAnalysisGoToDropdownRequest, _
                 ByVal dropdowns As IDropdownLists, _
                 ByVal formatter As ILLFormat, _
                 ByVal targetCell As Range)

    ValidateRequest request
    ValidateDropdowns dropdowns
    ValidateFormatter formatter
    ValidateTargetCell targetCell

    If Not request.HasEntries Then Exit Sub

    dropdowns.Add request.Values, request.ListName
    dropdowns.SetValidation targetCell, request.ListName

    formatter.ApplyFormat targetCell, LinelistSelectionDropdown

    targetCell.Value = request.LabelText
    targetCell.Name = request.CellName
End Sub

Private Sub ValidateRequest(ByVal request As IAnalysisGoToDropdownRequest)
    If request Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "request cannot be Nothing"
    End If
End Sub

Private Sub ValidateDropdowns(ByVal dropdowns As IDropdownLists)
    If dropdowns Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "dropdowns cannot be Nothing"
    End If
End Sub

Private Sub ValidateFormatter(ByVal formatter As ILLFormat)
    If formatter Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "formatter cannot be Nothing"
    End If
End Sub

Private Sub ValidateTargetCell(ByVal targetCell As Range)
    If targetCell Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "targetCell cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, TypeName(Me), messageText
End Sub
