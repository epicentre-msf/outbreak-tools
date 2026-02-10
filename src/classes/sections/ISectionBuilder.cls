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

'@class ISectionBuilder
'@description
'Defines the public contract for the SectionBuilder class, exposing a
'single Build method that iterates dictionary rows for a given sheet,
'groups them into sections and subsections, writes all variables via
'VarWriter, and applies section and subsection header formatting. The
'active mode (HList or VList) controls layout direction and companion
'sheet handling.

'@details
'Building modes control layout direction. HList builds horizontally on
'the main sheet plus printed and CRF companions. VList builds vertically
'on the main sheet only.
Public Enum SectionBuilderMode
    SectionBuilderModeHList = 1
    SectionBuilderModeVList = 2
End Enum


'@section Methods
'===============================================================================

'@jump:build
'@sub-title Build all sections, subsections, and variables for a sheet
'@details
'Scans the dictionary for rows matching sheetName starting from
'startRow, then iterates through sections and subsections, writing each
'variable via VarWriter and applying merged-cell section and subsection
'header formatting. For HList mode, headers appear on the main, printed,
'and CRF companion worksheets. For VList mode, headers appear on the
'main sheet only.
'@param sheetName String. The sheet name as it appears in the dictionary "sheet name" column.
'@param startRow Long. The first data row index in the dictionary for this sheet (1-based).
'@export
Public Sub Build(ByVal sheetName As String, ByVal startRow As Long)
End Sub
