VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Sections")
'@ModuleDescription("Interface for building sections, subsections, and variables on a linelist worksheet")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@description Building modes. HList builds on main + printed + CRF. VList builds on main only.
Public Enum SectionBuilderMode
    SectionBuilderModeHList = 1
    SectionBuilderModeVList = 2
End Enum

'@section Core Methods
'===============================================================================

'@description Build all sections, subsections, and variables for the given sheet.
'   Scans the dictionary for rows matching sheetName, then iterates sections,
'   subsections, and variables, applying formatting and writing variable data.
'@param sheetName The sheet name as it appears in the dictionary "sheet name" column.
'@param startRow The first row index in the dictionary for this sheet.
Public Sub Build(ByVal sheetName As String, ByVal startRow As Long)
End Sub
