VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetGraphSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable implementation of IAnalysisWorksheetGraphSpec capturing graph execution metadata")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisWorksheetGraphSpec

Public Enum AnalysisWorksheetGraphType
    AnalysisWorksheetGraphNone = 0
    AnalysisWorksheetGraphTimeSeries = 1
    AnalysisWorksheetGraphSpatioTemporal = 2
End Enum

Private Type TGraphSpecState
    graphType As AnalysisWorksheetGraphType
    tableId As String
    graphListId As String
    graphTitleListId As String
    goToPrefix As String
    goToSuffix As String
End Type

Private this As TGraphSpecState

Private Const CLASSNAME As String = "AnalysisWorksheetGraphSpec"

'@section Factory
'===============================================================================
'Create a graph specification describing how analysis graphs should be generated.
Public Function Create(ByVal graphType As AnalysisWorksheetGraphType, _
                       ByVal tableId As String, _
                       Optional ByVal graphListId As String = vbNullString, _
                       Optional ByVal graphTitleListId As String = vbNullString, _
                       Optional ByVal goToPrefix As String = vbNullString, _
                       Optional ByVal goToSuffix As String = vbNullString) As IAnalysisWorksheetGraphSpec

    ValidateGraphType graphType
    ValidateRequiredText tableId, "tableId"

    With New AnalysisWorksheetGraphSpec
        .Initialise graphType, tableId, graphListId, graphTitleListId, goToPrefix, goToSuffix
        Set Create = .Self
    End With
End Function

'Expose the instance through its interface type.
Public Property Get Self() As IAnalysisWorksheetGraphSpec
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal graphType As AnalysisWorksheetGraphType, _
                      ByVal tableId As String, _
                      ByVal graphListId As String, _
                      ByVal graphTitleListId As String, _
                      ByVal goToPrefix As String, _
                      ByVal goToSuffix As String)

    this.graphType = graphType
    this.tableId = Trim$(tableId)
    this.graphListId = Trim$(graphListId)
    this.graphTitleListId = Trim$(graphTitleListId)
    this.goToPrefix = Trim$(goToPrefix)
    this.goToSuffix = Trim$(goToSuffix)
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateGraphType(ByVal graphType As AnalysisWorksheetGraphType)
    Select Case graphType
        Case AnalysisWorksheetGraphNone, AnalysisWorksheetGraphTimeSeries, AnalysisWorksheetGraphSpatioTemporal
            'Valid values
        Case Else
            RaiseError CLng(ProjectError.InvalidArgument), "graphType is not recognised"
    End Select
End Sub

Private Sub ValidateRequiredText(ByVal candidate As String, ByVal argumentName As String)
    If LenB(Trim$(candidate)) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be empty"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisWorksheetGraphSpec_GraphType() As AnalysisWorksheetGraphType
    IAnalysisWorksheetGraphSpec_GraphType = this.graphType
End Property

Private Property Get IAnalysisWorksheetGraphSpec_TableId() As String
    IAnalysisWorksheetGraphSpec_TableId = this.tableId
End Property

Private Property Get IAnalysisWorksheetGraphSpec_GraphListId() As String
    IAnalysisWorksheetGraphSpec_GraphListId = this.graphListId
End Property

Private Property Get IAnalysisWorksheetGraphSpec_GraphTitleListId() As String
    IAnalysisWorksheetGraphSpec_GraphTitleListId = this.graphTitleListId
End Property

Private Property Get IAnalysisWorksheetGraphSpec_GoToPrefix() As String
    IAnalysisWorksheetGraphSpec_GoToPrefix = this.goToPrefix
End Property

Private Property Get IAnalysisWorksheetGraphSpec_GoToSuffix() As String
    IAnalysisWorksheetGraphSpec_GoToSuffix = this.goToSuffix
End Property
