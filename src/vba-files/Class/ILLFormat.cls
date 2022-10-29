VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface of linelist format")
'@IgnoreModule

Option Explicit

'Exposed methods
'Scope is where to apply the format: Those are the possibilities
Public Enum FormatScopes
  AnalysisSection = 1 '
  AnalysisCategoriesUnivariateRow = 2
  AnalysisPercent = 3
  AnalysisInteriorTable = 5
  AnalysisWholeTable = 6
  AnalysisUnivariateHeader = 7
  AnalysisCategoriesBivariateColumn = 8
  AnalysisTimeSeriesHeader = 9
  AnalysisMissingRow = 10
  AnalysisMissingColumn = 11
  AnalysisFirstRow = 12
  AnalysisOneCell = 14
  LinelistSelectionDropdown = 15
  VListSection = 16
  VListSubSection = 17
  HListHeader = 18
  HListGeo = 19
  LinelistShape = 20
  LinelistGeoShape = 21
  HListSection = 22
  HListSubSection = 23
  VListBorders = 24
  VListInnerBorders = 25
  HListBorders = 26
  HListCalculatedFormula = 27
  HListNote = 28
  HListTable = 29
  AnalysisTableTitle = 30
  AllAnalysisSheet = 31
  AnalysisCategoriesLabelRow = 32
  AnalysisCategoriesLabelColumn = 33
  AnalysisGlobalSummaryRow = 34
  AnalysisCategoriesBivariateRow = 35
  AnalysisOneCellLeft = 36
End Enum


Public Sub ApplyFormat(ByVal formattingObject As Variant, ByVal scope As FormatScopes)
End Sub

Public Function DesignValue(ByVal labelName As String, _
                            Optional ByVal returnedColor As Boolean = True) As Variant
End Function
