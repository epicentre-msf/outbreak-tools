VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListSectionDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Describes a logical section within a linelist worksheet")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Section identity
'===============================================================================
'Stable identifier used for dropdowns and navigation.
Public Property Get SectionName() As String
End Property

'Row anchoring point used by layout strategies.
Public Property Get AnchorRow() As Long
End Property

'Column anchoring point used by layout strategies.
Public Property Get AnchorColumn() As Long
End Property

