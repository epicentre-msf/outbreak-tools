VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface of linelist format")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
'Scope is where to apply the format: Those are the possibilities
Public Enum FormatScopes
  AdminLargeButton = 1
  AllAnalysisSheet = 2
  AllLinelistSheet = 3
  AnalysisCategoriesBivariateColumn = 4
  AnalysisCategoriesBivariateRow = 5
  AnalysisCategoriesLabelColumn = 6
  AnalysisCategoriesLabelRow = 7
  AnalysisCategoriesUnivariateRow  = 8
  AnalysisDateCell = 9
  AnalysisEndTable = 10
  AnalysisFirstRow = 11
  AnalysisGlobalSummaryRow  = 12
  AnalysisInfoCell = 13
  AnalysisInfoCellLeft = 14
  AnalysisInteriorTable = 15
  AnalysisMissingColumn = 16
  AnalysisMissingRow = 17
  AnalysisMissingRowWithoutHeader = 18
  AnalysisOneCell = 19
  AnalysisOneCellCenter = 20
  AnalysisOneCellLeft = 21
  AnalysisPercent = 22
  AnalysisSection = 23
  AnalysisTableTitle = 24
  AnalysisTimeSeriesHeader = 25
  AnalysisUnivariateHeader = 26
  AnalysisWholeTable = 27
  CRFChoiceHeader = 28
  CRFChoiceValues = 29
  CRFDate = 30
  CRFDecimal = 31
  CRFFormatCustom = 32
  CRFInteger = 33
  CRFPercentage = 34
  CRFText = 35
  HListBorders = 36
  HListCalculatedFormulaCell = 37
  HListCalculatedFormulaHeader = 38
  HListCRFSection = 39
  HListCRFSheet = 40
  HListCRFSubSection = 41
  HListCustomPivotTables = 42
  HListGeo = 43
  HListGeoHeader = 44
  HListHeader = 45
  HListMainLab = 46
  HListNote = 47
  HListPrintedHeader = 48
  HListPrintedSheet = 49
  HListPrintedTable = 50
  HListSection = 51
  HListSimpleHeaderCell = 52
  HListSublab = 53
  HListSubSection = 54
  HListTable = 55
  LinelistGeoButton = 56
  LinelistGeoShape = 57
  LinelistHiddenCell  = 58
  LinelistSelectionDropdown = 59
  LinelistShape = 60
  LinelistSmallButton = 61
  LinelistWarning = 62
  VListBorders = 63
  VListInnerBorders = 64
  VListMainLab = 65
  VListSection = 66
  VListSublab = 67
  VListSubSection = 68
End Enum

Public Sub ApplyFormat(ByVal formattingObject As Variant, ByVal scope As FormatScopes, _ 
                       Optional ByVal nbDecimals As Integer = 0)
End Sub

Public Function DesignValue(ByVal labelName As String, _
                            Optional ByVal returnedColor As Boolean = True) As Variant
End Function

Public Sub Import(ByVal inpsh As Worksheet)
End Sub
