VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetPipeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Iterates worksheet specifications and delegates execution to a handler")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisWorksheetPipeline

Private Const CLASSNAME As String = "AnalysisWorksheetPipeline"

'@section Validation helpers
'===============================================================================
Private Sub ValidateSequence(ByVal sequence As BetterArray)
    If sequence Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "sequence cannot be Nothing"
    End If
End Sub

Private Sub ValidateHandler(ByVal handler As IAnalysisWorksheetHandler)
    If handler Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "handler cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpecification(ByVal spec As IAnalysisWorksheetSpec, ByVal index As Long)
    If spec Is Nothing Then
        RaiseError CLng(ProjectError.InvalidState), "Worksheet specification at index " & CStr(index) & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisWorksheetPipeline_Run(ByVal sequence As BetterArray, ByVal handler As IAnalysisWorksheetHandler)
    Dim idx As Long
    Dim spec As IAnalysisWorksheetSpec

    ValidateSequence sequence
    ValidateHandler handler

    handler.BeginSequence

    If sequence.Length > 0 Then
        For idx = sequence.LowerBound To sequence.UpperBound
            Set spec = sequence.Item(idx)
            ValidateSpecification spec, idx
            handler.ProcessWorksheet spec
            Set spec = Nothing
        Next idx
    End If

    handler.CompleteSequence
End Sub
