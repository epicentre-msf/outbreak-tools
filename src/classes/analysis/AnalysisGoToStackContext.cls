VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToStackContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Context object accumulating section and header labels during plan execution")
'@depends ProjectError, BetterArray, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TStackContextState
    sections As BetterArray
    headers As BetterArray
End Type

Private this As TStackContextState
Private Const CLASSNAME As String = "AnalysisGoToStackContext"

Private Sub Class_Initialize()
    Set this.sections = AnalysisArrayUtilities.CreateBuffer()
    Set this.headers = AnalysisArrayUtilities.CreateBuffer()
End Sub

'@section Public API
'===============================================================================
Public Sub Reset()
    this.sections.Clear
    this.headers.Clear
End Sub

Public Sub AddSectionEntry(ByVal entry As IAnalysisGoToEntry)
    ValidateEntry entry, "entry"
    this.sections.Push entry
End Sub

Public Sub AddHeaderEntry(ByVal entry As IAnalysisGoToEntry)
    ValidateEntry entry, "entry"
    this.headers.Push entry
End Sub

Public Property Get SectionLabels() As BetterArray
    Set SectionLabels = this.sections
End Property

Public Property Get HeaderLabels() As BetterArray
    Set HeaderLabels = this.headers
End Property

Private Sub ValidateEntry(ByVal entry As IAnalysisGoToEntry, ByVal argumentName As String)
    If entry Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub
