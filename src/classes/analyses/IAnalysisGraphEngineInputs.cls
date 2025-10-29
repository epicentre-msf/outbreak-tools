VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphEngineInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Provides graph-related dependencies required by AnalysisGraphEngine")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the cross table associated with the supplied plan item.
Public Function CrossTableFor(ByVal planItem As IAnalysisGraphPlanItem) As ICrossTable
End Function

'Return the ListObject containing graph definitions for time-series analysis.
Public Property Get GraphList() As ListObject
End Property

'Return the ListObject containing time-series configuration rows.
Public Property Get TimeSeriesList() As ListObject
End Property

'Return the ListObject containing graph title metadata.
Public Property Get TitleList() As ListObject
End Property

'Return the linelist specifications required to evaluate graph sources.
Public Property Get Linelist() As ILinelistSpecs
End Property
