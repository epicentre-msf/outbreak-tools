VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of linelist format"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of linelist format")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILLFormat
'@description
'Defines the public contract for the LLFormat class, exposing
'design-driven formatting of worksheets, ranges, shapes, and other
'Excel objects. Reads style metadata from a format table and applies
'it according to FormatScope constants. Supports import/export of
'format definitions across workbooks.


'@section Enumerations
'===============================================================================
'@description Supported formatting scopes. Each constant identifies a
'specific visual pattern that ApplyFormat applies to the supplied object.

Public Enum FormatScope
  AdminLargeButton = 1
  AllAnalysisSheet = 2
  AllLinelistSheet = 3
  AnalysisCategoriesBivariateColumn = 4
  AnalysisCategoriesBivariateRow = 5
  AnalysisCategoriesLabelColumn = 6
  AnalysisCategoriesLabelRow = 7
  AnalysisCategoriesUnivariateRow = 8
  AnalysisDateCell = 9
  AnalysisEndTable = 10
  AnalysisFirstRow = 11
  AnalysisGlobalSummaryRow = 12
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
  LinelistHiddenCell = 58
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


'@section Formatting
'===============================================================================
'@description Core formatting operation.

'@jump:apply-format
'@sub-title Apply a design-driven format to an Excel object
'@details
'Reads style metadata from the format table and applies it to the
'supplied object (Worksheet, Range, Shape, ListObject, or Characters)
'according to the given FormatScope constant.
'@param formattingObject Variant. The Excel object to format.
'@param scope Byte. FormatScope constant selecting the visual pattern.
'@param nbDecimals Integer. Decimal count for numeric formatting. Defaults to 0.
Public Sub ApplyFormat(formattingObject As Variant, ByVal scope As Byte, _
                       Optional ByVal nbDecimals As Integer = 0)
End Sub


'@section Design Lookups
'===============================================================================
'@description Retrieve raw style values from the format table.

'@jump:design-value
'@sub-title Look up a design value or colour from the format table
'@details
'Retrieves either the fill colour (Interior.Color) or the raw cell
'value for the specified label row under the active design column.
'@param labelName String. Label text identifying the row in the format table.
'@param returnedColor Boolean. When True returns Interior.Color; otherwise returns the raw value. Defaults to True.
'@return Variant. The colour (Long) or cell value.
Public Function DesignValue(ByVal labelName As String, _
                            Optional ByVal returnedColor As Boolean = True) As Variant
End Function


'@section DataExchange
'===============================================================================
'@description Import and export format definitions.

'@jump:import
'@sub-title Import formatting definitions from another worksheet
'@details
'Reads design columns from the source worksheet and merges them into
'the current format table. Updates the active design name from the
'source DESIGNTYPE named range.
'@param inpsh Worksheet. Source worksheet containing the format table.
Public Sub Import(ByVal inpsh As Worksheet)
End Sub

'@jump:export
'@sub-title Export the format table to another workbook
'@details
'Copies the format table and cell styles to a new worksheet in the
'destination workbook. When the target worksheet already exists,
'imports from it instead.
'@param toWkb Workbook. Destination workbook.
Public Sub Export(ByVal toWkb As Workbook)
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the format object has logged diagnostic messages
'@details
'Returns True when at least one diagnostic entry has been recorded,
'typically for missing labels in the format table.
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean
End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal IChecking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return IChecking. The checking instance, or Nothing.
Public Property Get CheckingValues() As IChecking
End Property
