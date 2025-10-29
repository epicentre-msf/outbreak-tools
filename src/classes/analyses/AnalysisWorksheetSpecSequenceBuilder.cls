VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetSpecSequenceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Collects AnalysisWorksheetSpec instances into an ordered sequence")
'@depends ProjectError, AnalysisWorksheetSpec, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisWorksheetSpecSequenceBuilder

Private Type TSequenceBuilderState
    specs As BetterArray
End Type

Private this As TSequenceBuilderState

Private Const CLASSNAME As String = "AnalysisWorksheetSpecSequenceBuilder"

Private Sub Class_Initialize()
    Set this.specs = AnalysisArrayUtilities.CreateBuffer()
End Sub

'@section Helper factories
'===============================================================================
Private Sub ValidateSpec(ByVal spec As IAnalysisWorksheetSpec)
    If spec Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "spec cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Sub AddInternal(ByVal spec As IAnalysisWorksheetSpec)
    ValidateSpec spec
    this.specs.Push spec
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisWorksheetSpecSequenceBuilder_Add(ByVal spec As IAnalysisWorksheetSpec)
    AddInternal spec
End Sub

Private Sub IAnalysisWorksheetSpecSequenceBuilder_AddSpec(ByVal scope As AnalysisScope, _
                                                          ByVal tableIds As Variant, _
                                                          Optional ByVal tableTypeCodes As Variant = Empty, _
                                                          Optional ByVal graphSpecs As BetterArray = Nothing, _
                                                          Optional ByVal sectionGoToPrefix As String = vbNullString, _
                                                          Optional ByVal graphGoToPrefix As String = vbNullString, _
                                                          Optional ByVal graphGoToSuffix As String = vbNullString)
    Dim worksheetSpec As IAnalysisWorksheetSpec
    Set worksheetSpec = AnalysisWorksheetSpec.Create(scope, tableIds, tableTypeCodes, graphSpecs, sectionGoToPrefix, graphGoToPrefix, graphGoToSuffix)
    AddInternal worksheetSpec
End Sub

Private Function IAnalysisWorksheetSpecSequenceBuilder_Build() As BetterArray
    Set IAnalysisWorksheetSpecSequenceBuilder_Build = AnalysisArrayUtilities.CloneItems(this.specs)
End Function

Private Property Get IAnalysisWorksheetSpecSequenceBuilder_Count() As Long
    IAnalysisWorksheetSpecSequenceBuilder_Count = this.specs.Length
End Property
