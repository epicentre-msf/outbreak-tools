VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VerticalListLayoutStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Concrete IListLayoutStrategy for vertical (VList) sheets")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListLayoutStrategy

Private Type TLayoutState
    processedSections As BetterArray
    initialiseCount As Long
    finaliseCount As Long
End Type

Private this As TLayoutState
Private Const CLASSNAME As String = "VerticalListLayoutStrategy"

'@section Diagnostics
'===============================================================================
Public Property Get ProcessedSections() As BetterArray
    If this.processedSections Is Nothing Then
        Set ProcessedSections = New BetterArray
        ProcessedSections.LowerBound = 1
    Else
        Set ProcessedSections = this.processedSections.Clone()
    End If
End Property

Public Property Get InitialiseCount() As Long
    InitialiseCount = this.initialiseCount
End Property

Public Property Get FinaliseCount() As Long
    FinaliseCount = this.finaliseCount
End Property

'@section Helpers
'===============================================================================
Private Sub EnsureProcessedSections()
    If this.processedSections Is Nothing Then
        Set this.processedSections = New BetterArray
        this.processedSections.LowerBound = 1
    End If
End Sub

Private Sub ValidateSection(ByVal section As IListSectionDescriptor)
    If section Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "section cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IListLayoutStrategy_Initialise(ByVal context As IListBuildContext)
    this.initialiseCount = this.initialiseCount + 1
    EnsureProcessedSections
End Sub

Private Sub IListLayoutStrategy_ApplySection(ByVal context As IListBuildContext, _
                                             ByVal section As IListSectionDescriptor)
    EnsureProcessedSections
    ValidateSection section
    this.processedSections.Push section
End Sub

Private Sub IListLayoutStrategy_Finalise(ByVal context As IListBuildContext)
    this.finaliseCount = this.finaliseCount + 1
End Sub

Private Sub IListLayoutStrategy_Invalidate()
    Set this.processedSections = Nothing
End Sub

