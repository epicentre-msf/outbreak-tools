VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetSpecSequenceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Builds ordered collections of AnalysisWorksheetSpec instances")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Add a pre-built worksheet specification to the builder.
Public Sub Add(ByVal spec As IAnalysisWorksheetSpec)
End Sub

'Convenience helper to build and add a worksheet specification from simple parameters.
Public Sub AddSpec(ByVal scope As AnalysisScope, _
                   ByVal tableIds As Variant, _
                   Optional ByVal tableTypeCodes As Variant = Empty, _
                   Optional ByVal graphSpecs As BetterArray = Nothing, _
                   Optional ByVal sectionGoToPrefix As String = vbNullString, _
                   Optional ByVal graphGoToPrefix As String = vbNullString, _
                   Optional ByVal graphGoToSuffix As String = vbNullString)
End Sub

'Return the ordered list of worksheet specifications.
Public Function Build() As BetterArray
End Function

'Return the number of worksheet specifications currently stored.
Public Property Get Count() As Long
End Property
