VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGoToDropdownRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Represents a dropdown request derived from GoTo entries")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the dropdown scope (section/header/graph).
Public Property Get Scope() As String
End Property

'Return the list name registered in IDropdownLists.
Public Property Get ListName() As String
End Property

'Return the cell name assigned to the target range.
Public Property Get CellName() As String
End Property

'Return the label shown to the user in the cell.
Public Property Get LabelText() As String
End Property

'Return the underlying GoTo entries.
Public Property Get Entries() As BetterArray
End Property

'Return the dropdown values (display text) cloned for safety.
Public Property Get Values() As BetterArray
End Property

'Indicates whether the request contains entries to apply.
Public Property Get HasEntries() As Boolean
End Property
