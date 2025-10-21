VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable typed representation of an Analysis GoTo dropdown entry")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGoToEntry

Private Type TGoToEntryState
    scope As String
    labelText As String
    prefixText As String
    suffixText As String
    displayText As String
End Type

Private this As TGoToEntryState

Private Const CLASSNAME As String = "AnalysisGoToEntry"

'@section Factory
'===============================================================================
'Create a typed GoTo entry ensuring mandatory fields are validated.
Public Function Create(ByVal scope As String, _
                       ByVal labelText As String, _
                       Optional ByVal prefixText As String = vbNullString, _
                       Optional ByVal suffixText As String = vbNullString) As IAnalysisGoToEntry

    ValidateScope scope
    ValidateLabel labelText

    With New AnalysisGoToEntry
        .Initialise scope, labelText, prefixText, suffixText
        Set Create = .Self
    End With
End Function

'Return the instance as its interface type.
Public Property Get Self() As IAnalysisGoToEntry
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal scope As String, _
                      ByVal labelText As String, _
                      ByVal prefixText As String, _
                      ByVal suffixText As String)

    this.scope = Normalize(scope)
    this.labelText = Normalize(labelText)
    this.prefixText = Normalize(prefixText)
    this.suffixText = Trim$(suffixText)
    this.displayText = BuildDisplayText(this.prefixText, this.labelText)
End Sub

Private Function BuildDisplayText(ByVal prefixText As String, ByVal labelText As String) As String
    If LenB(prefixText) = 0 Then
        BuildDisplayText = labelText
    Else
        BuildDisplayText = prefixText & labelText
    End If
End Function

Private Function Normalize(ByVal valueText As String) As String
    Normalize = Trim$(valueText)
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateScope(ByVal scope As String)
    If Normalize(scope) = vbNullString Then
        RaiseError CLng(ProjectError.InvalidArgument), "scope cannot be empty"
    End If
End Sub

Private Sub ValidateLabel(ByVal labelText As String)
    If Normalize(labelText) = vbNullString Then
        RaiseError CLng(ProjectError.InvalidArgument), "labelText cannot be empty"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGoToEntry_Scope() As String
    IAnalysisGoToEntry_Scope = this.scope
End Property

Private Property Get IAnalysisGoToEntry_Label() As String
    IAnalysisGoToEntry_Label = this.labelText
End Property

Private Property Get IAnalysisGoToEntry_Prefix() As String
    IAnalysisGoToEntry_Prefix = this.prefixText
End Property

Private Property Get IAnalysisGoToEntry_Suffix() As String
    IAnalysisGoToEntry_Suffix = this.suffixText
End Property

Private Property Get IAnalysisGoToEntry_DisplayText() As String
    IAnalysisGoToEntry_DisplayText = this.displayText
End Property
