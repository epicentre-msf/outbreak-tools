VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable implementation of IAnalysisWorksheetSpec describing worksheet orchestration metadata")
'@depends ProjectError, AnalysisWorksheetGraphSpec, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisWorksheetSpec

Private Type TWorksheetSpecState
    scope As AnalysisScope
    tableIds As BetterArray
    tableTypeCodes As BetterArray
    graphSpecs As BetterArray
    sectionPrefix As String
    graphPrefix As String
    graphSuffix As String
End Type

Private this As TWorksheetSpecState

Private Const CLASSNAME As String = "AnalysisWorksheetSpec"

'@section Factory
'===============================================================================
'Create a worksheet specification using the supplied metadata.
Public Function Create(ByVal scope As AnalysisScope, _
                       ByVal tableIds As Variant, _
                       Optional ByVal tableTypeCodes As Variant = Empty, _
                       Optional ByVal graphSpecs As BetterArray = Nothing, _
                       Optional ByVal sectionGoToPrefix As String = vbNullString, _
                       Optional ByVal graphGoToPrefix As String = vbNullString, _
                       Optional ByVal graphGoToSuffix As String = vbNullString) As IAnalysisWorksheetSpec

    With New AnalysisWorksheetSpec
        .Initialise scope, tableIds, tableTypeCodes, graphSpecs, sectionGoToPrefix, graphGoToPrefix, graphGoToSuffix
        Set Create = .Self
    End With
End Function

'Expose the instance through its interface.
Public Property Get Self() As IAnalysisWorksheetSpec
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal scope As AnalysisScope, _
                      ByVal tableIds As Variant, _
                      ByVal tableTypeCodes As Variant, _
                      ByVal graphSpecs As BetterArray, _
                      ByVal sectionGoToPrefix As String, _
                      ByVal graphGoToPrefix As String, _
                      ByVal graphGoToSuffix As String)

    this.scope = scope
    Set this.tableIds = NormalizeTableIds(tableIds)
    Set this.tableTypeCodes = NormalizeTableTypes(tableTypeCodes, this.tableIds.Length)
    Set this.graphSpecs = NormalizeGraphSpecs(graphSpecs)
    this.sectionPrefix = Trim$(sectionGoToPrefix)
    this.graphPrefix = Trim$(graphGoToPrefix)
    this.graphSuffix = Trim$(graphGoToSuffix)
End Sub

'@section Normalisation helpers
'===============================================================================
Private Function NormalizeTableIds(ByVal tableIds As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim item As Variant
    Dim stringValue As String

    Set result = AnalysisArrayUtilities.CreateBuffer()

    If IsObject(tableIds) Then
        If TypeName(tableIds) = "BetterArray" Then
            Dim source As BetterArray
            Set source = tableIds
            If Not source Is Nothing Then
                If source.Length > 0 Then
                    For idx = source.LowerBound To source.UpperBound
                        stringValue = Trim$(CStr(source.Item(idx)))
                        If LenB(stringValue) > 0 Then result.Push stringValue
                    Next idx
                End If
            End If
        Else
            RaiseError CLng(ProjectError.InvalidArgument), "tableIds must be an array of strings or BetterArray"
        End If
    ElseIf IsArray(tableIds) Then
        For idx = LBound(tableIds) To UBound(tableIds)
            item = tableIds(idx)
            stringValue = Trim$(CStr(item))
            If LenB(stringValue) = 0 Then GoTo ContinueLoop
            result.Push stringValue
ContinueLoop:
        Next idx
    Else
        stringValue = Trim$(CStr(tableIds))
        If LenB(stringValue) > 0 Then result.Push stringValue
    End If

    If result.Length = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "tableIds must contain at least one entry"
    End If

    Set NormalizeTableIds = result
End Function

Private Function NormalizeGraphSpecs(ByVal graphSpecs As BetterArray) As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim graphSpec As IAnalysisWorksheetGraphSpec

    Set result = AnalysisArrayUtilities.CreateBuffer()

    If graphSpecs Is Nothing Then
        Set NormalizeGraphSpecs = result
        Exit Function
    End If

    If graphSpecs.Length = 0 Then
        Set NormalizeGraphSpecs = result
        Exit Function
    End If

    For idx = graphSpecs.LowerBound To graphSpecs.UpperBound
        Set graphSpec = graphSpecs.Item(idx)
        ValidateGraphSpec graphSpec
        result.Push graphSpec
    Next idx

    Set NormalizeGraphSpecs = result
End Function

Private Function NormalizeTableTypes(ByVal tableTypeCodes As Variant, ByVal expectedCount As Long) As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim value As Variant

    Set result = AnalysisArrayUtilities.CreateBuffer()

    If IsEmpty(tableTypeCodes) Then
        For idx = 1 To expectedCount
            result.Push CByte(0)
        Next idx
        Set NormalizeTableTypes = result
        Exit Function
    End If

    If IsArray(tableTypeCodes) Then
        For idx = LBound(tableTypeCodes) To UBound(tableTypeCodes)
            value = tableTypeCodes(idx)
            result.Push CByte(value)
        Next idx
    ElseIf TypeName(tableTypeCodes) = "BetterArray" Then
        Dim source As BetterArray
        Set source = tableTypeCodes
        If Not source Is Nothing Then
            If source.Length > 0 Then
                For idx = source.LowerBound To source.UpperBound
                    result.Push CByte(source.Item(idx))
                Next idx
            End If
        End If
    Else
        result.Push CByte(tableTypeCodes)
    End If

    If result.Length <> expectedCount Then
        RaiseError CLng(ProjectError.InvalidArgument), "tableTypeCodes length must match tableIds"
    End If

    Set NormalizeTableTypes = result
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateGraphSpec(ByVal graphSpec As IAnalysisWorksheetGraphSpec)
    If graphSpec Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "graphSpec cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Function CloneTableIds() As BetterArray
    Set CloneTableIds = AnalysisArrayUtilities.CloneItems(this.tableIds)
End Function

Private Function CloneGraphSpecs() As BetterArray
    Set CloneGraphSpecs = AnalysisArrayUtilities.CloneItems(this.graphSpecs)
End Function

Private Function CloneTableTypeCodes() As BetterArray
    Set CloneTableTypeCodes = AnalysisArrayUtilities.CloneItems(this.tableTypeCodes)
End Function

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisWorksheetSpec_Scope() As AnalysisScope
    IAnalysisWorksheetSpec_Scope = this.scope
End Property

Private Property Get IAnalysisWorksheetSpec_TableIds() As BetterArray
    Set IAnalysisWorksheetSpec_TableIds = CloneTableIds()
End Property

Private Property Get IAnalysisWorksheetSpec_TableTypeCodes() As BetterArray
    Set IAnalysisWorksheetSpec_TableTypeCodes = CloneTableTypeCodes()
End Property

Private Property Get IAnalysisWorksheetSpec_GraphSpecs() As BetterArray
    Set IAnalysisWorksheetSpec_GraphSpecs = CloneGraphSpecs()
End Property

Private Property Get IAnalysisWorksheetSpec_SectionGoToPrefix() As String
    IAnalysisWorksheetSpec_SectionGoToPrefix = this.sectionPrefix
End Property

Private Property Get IAnalysisWorksheetSpec_GraphGoToPrefix() As String
    IAnalysisWorksheetSpec_GraphGoToPrefix = this.graphPrefix
End Property

Private Property Get IAnalysisWorksheetSpec_GraphGoToSuffix() As String
    IAnalysisWorksheetSpec_GraphGoToSuffix = this.graphSuffix
End Property

Private Property Get IAnalysisWorksheetSpec_HasGraphs() As Boolean
    If this.graphSpecs Is Nothing Then
        IAnalysisWorksheetSpec_HasGraphs = False
    Else
        IAnalysisWorksheetSpec_HasGraphs = (this.graphSpecs.Length > 0)
    End If
End Property
