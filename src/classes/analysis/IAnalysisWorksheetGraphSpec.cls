VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetGraphSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Represents a graph generation request within an AnalysisWorksheetSpec")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the analysis graph type encoded as AnalysisWorksheetGraphType.
Public Property Get GraphType() As AnalysisWorksheetGraphType
End Property

'Return the primary table identifier associated with the graph request.
Public Property Get TableId() As String
End Property

'Return the ListObject identifier providing graph metadata (time-series lists, etc.).
Public Property Get GraphListId() As String
End Property

'Return the ListObject identifier supplying graph titles (time-series scenarios).
Public Property Get GraphTitleListId() As String
End Property

'Return the GoTo prefix that should be applied when registering graph navigation entries.
Public Property Get GoToPrefix() As String
End Property

'Return the GoTo suffix that should be applied when registering graph navigation entries.
Public Property Get GoToSuffix() As String
End Property
