VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableSequenceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Builds ordered policy-evaluated table specifications for analysis output scopes")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Enumerate the specifications associated with a listobject and evaluate their policies.
Public Function Build(ByVal specificationList As ListObject, _
                      ByVal linelistSpecs As ILinelistSpecs, _
                      ByVal tableTypeCode As Byte) As BetterArray
End Function
