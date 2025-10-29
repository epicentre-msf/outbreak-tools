VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGoToEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Represents a typed entry destined for Analysis GoTo dropdowns")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the scope associated with the entry (e.g., section, header, graph).
Public Property Get Scope() As String
End Property

'Return the trimmed label text associated with the dropdown entry.
Public Property Get Label() As String
End Property

'Return the prefix applied to the label when building the display value.
Public Property Get Prefix() As String
End Property

'Return the optional suffix bound to the entry (typically a table or section identifier).
Public Property Get Suffix() As String
End Property

'Return the display text that should be surfaced in the dropdown control.
Public Property Get DisplayText() As String
End Property
