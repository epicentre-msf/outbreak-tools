VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTablePlanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Aggregates the planned table items and navigation metadata for analysis output")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the ordered plan items to process.
Public Property Get Items() As BetterArray
End Property

'Return the section navigation labels generated during planning.
Public Property Get SectionLabels() As BetterArray
End Property

'Return the header navigation labels generated during planning.
Public Property Get HeaderLabels() As BetterArray
End Property

'Return the prefix applied to section labels.
Public Property Get SectionPrefix() As String
End Property

'Return the prefix applied to header labels.
Public Property Get HeaderPrefix() As String
End Property
