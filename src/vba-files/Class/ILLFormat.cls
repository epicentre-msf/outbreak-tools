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
  AllAnalysisSheet = 1
  AnalysisCategoriesBivariateColumn = 2
  AnalysisCategoriesBivariateRow = 3
  AnalysisCategoriesLabelColumn = 4
  AnalysisCategoriesLabelRow = 5
  AnalysisCategoriesUnivariateRow = 6
  AnalysisFirstRow = 7
  AnalysisGlobalSummaryRow = 8
  LinelistHiddenCell = 9
  AnalysisInteriorTable = 10
  AnalysisMissingColumn = 11
  AnalysisMissingRow = 12
  AnalysisOneCell = 13
  AnalysisOneCellLeft = 14
  AnalysisPercent = 15
  AnalysisSection = 16
  AnalysisTableTitle = 17
  AnalysisTimeSeriesHeader = 18
  AnalysisUnivariateHeader = 19
  AnalysisWholeTable = 20
  HListBorders = 21
  HListCalculatedFormulaCell = 22
  HListGeo = 23
  HListHeader = 24
  HListNote = 25
  HListSection = 26
  HListSubSection = 27
  HListTable = 28
  LinelistGeoShape = 29
  LinelistSelectionDropdown = 30
  LinelistShape = 31
  VListBorders = 32
  VListInnerBorders = 33
  VListSection = 34
  VListSubSection = 35
  AnalysisEndTable = 36
  AnalysisOneCellCenter = 37
  AnalysisInfoCell = 38
  HListCalculatedFormulaHeader = 39
  HListSublab = 40
  HListMainLab = 41
  VListMainLab = 42
  VListSublab = 43
  HListGeoHeader = 44
  LinelistSmallButton = 45
  LinelistGeoButton = 46
  AdminLargeButton = 47
End Enum

Public Sub ApplyFormat(ByVal formattingObject As Variant, ByVal scope As FormatScopes)
End Sub

Public Function DesignValue(ByVal labelName As String, _
                            Optional ByVal returnedColor As Boolean = True) As Variant
End Function


