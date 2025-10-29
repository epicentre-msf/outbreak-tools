VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Contract for classes able to enumerate analysis table specifications")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Enumerate table specifications contained in the supplied listobject.
Public Function Enumerate(ByVal specificationList As ListObject, _
                          ByVal linelistSpecs As ILinelistSpecs, _
                          ByVal tableTypeCode As Byte) As BetterArray
End Function

