VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Defines the worksheet-level instructions required to run analysis output pipelines")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the worksheet scope being orchestrated.
Public Property Get Scope() As AnalysisScope
End Property


'Return the ordered list of ListObject identifiers to render as tables.
Public Property Get TableIds() As BetterArray
End Property

'Return the corresponding table type codes (AnalysisTablesType) aligned with TableIds.
Public Property Get TableTypeCodes() As BetterArray
End Property

'Return the graph specifications associated with the worksheet.
Public Property Get GraphSpecs() As BetterArray
End Property

'Return the GoTo prefix used when registering section dropdown entries.
Public Property Get SectionGoToPrefix() As String
End Property

'Return the GoTo prefix applied to graph dropdown entries.
Public Property Get GraphGoToPrefix() As String
End Property

'Return the GoTo suffix applied to graph dropdown entries.
Public Property Get GraphGoToSuffix() As String
End Property

'Indicates whether the worksheet has graphs configured.
Public Property Get HasGraphs() As Boolean
End Property
