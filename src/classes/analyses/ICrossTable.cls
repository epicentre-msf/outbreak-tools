VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for cross-table building"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for cross-table building")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ICrossTable
'@description
'Defines the public contract for the CrossTable class. Exposes properties
'for table geometry (rows, columns, ranges) and methods for building,
'formatting, and naming the cross-table on an analysis output worksheet.


'@section Properties
'===============================================================================
'@description Read-only accessors for table geometry, navigation, and metadata.

'@jump:header-range
'@prop-title Header range of the table on the output worksheet
'@details
'Returns the Range spanning the column headers of the built cross-table.
'Used by formatting and formula-writing routines to locate header cells.
'@return Range. The header row range.
Public Property Get HeaderRange() As Range: End Property

'@jump:rows-categories-range
'@prop-title Row categories range with optional filtering
'@details
'Returns the Range spanning the row category labels. When includeHeaders
'is True, the header row is included. When onlyCategories is True, only
'the category label column is returned without accompanying data columns.
'@param includeHeaders Optional Boolean. When True, includes the header
'row in the returned range. Defaults to True.
'@param onlyCategories Optional Boolean. When True, returns only the
'label column. Defaults to False.
'@return Range. The row categories range.
Public Property Get RowsCategoriesRange( _
    Optional ByVal includeHeaders As Boolean = True, _
    Optional ByVal onlyCategories As Boolean = False) As Range
End Property

'@jump:column-range
'@prop-title Column data range for a specific column value
'@details
'Returns the Range spanning a single data column identified by its
'header value. Used by formula writers to target a specific column
'when building COUNTIFS or SUMPRODUCT expressions.
'@param colVal String. The column header value to locate.
'@param onlyCategories Optional Boolean. When True, returns only the
'category cells. Defaults to False.
'@param includeHeaders Optional Boolean. When True, includes the header
'cell. Defaults to False.
'@return Range. The data range for the specified column.
Public Property Get ColumnRange(ByVal colVal As String, _
    Optional ByVal onlyCategories As Boolean = False, _
    Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'@jump:row-range
'@prop-title Row data range for a specific row value
'@details
'Returns the Range spanning a single data row identified by its row
'label value. Used by formula writers to target a specific row when
'building cross-table expressions.
'@param rowVal String. The row label value to locate.
'@param includeHeaders Optional Boolean. When True, includes the row
'label cell. Defaults to True.
'@return Range. The data range for the specified row.
Public Property Get RowRange(ByVal rowVal As String, _
    Optional ByVal includeHeaders As Boolean = True) As Range
End Property

'@jump:start-row
'@prop-title Starting row of the table on the output worksheet
'@details
'Returns the 1-based row index where this cross-table begins. Computed
'from the previous table end row or the section start row when this is
'the first table.
'@return Long. The starting row index.
Public Property Get StartRow() As Long: End Property

'@jump:previous
'@prop-title Previous cross-table in the same section
'@details
'Returns the ICrossTable instance that precedes this one in the build
'chain. Returns Nothing when this is the first table in its section.
'Used to compute start-row offsets for stacked tables.
'@return ICrossTable. The preceding cross-table, or Nothing.
Public Property Get Previous() As ICrossTable: End Property

'@jump:specifications
'@prop-title Table specifications used to build this cross-table
'@details
'Returns the ITableSpecs instance that defines the configuration for
'this cross-table. Provides access to table type, variable names,
'category lists, and validation flags.
'@return ITableSpecs. The backing table specification.
Public Property Get Specifications() As ITableSpecs: End Property

'@jump:wksh
'@prop-title Output worksheet where the table is built
'@details
'Returns the Worksheet reference where this cross-table writes its
'headers, row labels, column labels, and data cells.
'@return Worksheet. The output worksheet.
Public Property Get Wksh() As Worksheet: End Property

'@jump:number-of-columns
'@prop-title Number of data columns in the table
'@details
'Returns the count of data columns excluding the row-label column.
'For bivariate and time series tables this equals the number of column
'categories plus optional total and missing columns.
'@return Long. The data column count.
Public Property Get NumberOfColumns() As Long: End Property

'@jump:end-column
'@prop-title Last column used by the table on the worksheet
'@details
'Returns the 1-based column index of the rightmost cell occupied by
'this cross-table. Used by sibling tables and formatting routines to
'avoid overlapping output.
'@return Long. The ending column index.
Public Property Get EndColumn() As Long: End Property

'@jump:end-row
'@prop-title Last row used by the table on the worksheet
'@details
'Returns the 1-based row index of the bottommost cell occupied by this
'cross-table. Used to compute the start row for the next stacked table.
'@return Long. The ending row index.
Public Property Get EndRow() As Long: End Property

'@jump:time-series-header
'@prop-title Header label for time series tables
'@details
'Returns the translated header text used for the time-period column in
'time series cross-tables. Returns an empty string for non-time-series
'table types.
'@return String. The time series column header label.
Public Property Get TimeSeriesHeader() As String: End Property

'@jump:named-ranges-list
'@prop-title List of named range names created by this table
'@details
'Returns a BetterArray containing all the named range names that
'NameRanges created for this cross-table. Used by AnaTabIds to register
'the ranges for export tracking.
'@return BetterArray. The collection of named range names.
Public Property Get NamedRangesList() As BetterArray: End Property


'@section Methods
'===============================================================================
'@description Operations for building and formatting the cross-table.

'@jump:addrows
'@sub-title Add row headers and row category labels to the worksheet
'@details
'Writes the row variable label and all row category values into the
'first column of the cross-table region. Includes optional total and
'missing rows based on the table specification flags.
'@export
Public Sub AddRows(): End Sub

'@jump:addcolumns
'@sub-title Add column headers and column category labels to the worksheet
'@details
'Writes the column variable label and all column category values into the
'header row of the cross-table region. Includes optional total and
'missing columns based on the table specification flags.
'@export
Public Sub AddColumns(): End Sub

'@jump:addheader
'@sub-title Add the table title and header to the worksheet
'@details
'Writes the table title cell above the cross-table and any section
'header when IsNewSection is True. Positions the header based on the
'start row computed from the previous table.
'@export
Public Sub AddHeader(): End Sub

'@jump:name-ranges
'@sub-title Create all named ranges for the table data regions
'@details
'Defines workbook-level named ranges for the row categories, column
'headers, data body, total row, total column, and other structural
'regions. The names follow a convention based on the table identifier
'and are registered in NamedRangesList for export tracking.
'@export
Public Sub NameRanges(): End Sub

'@jump:build
'@sub-title Build the complete cross-table on the worksheet
'@details
'Orchestrates the full table construction by calling AddHeader, AddRows,
'AddColumns, and NameRanges in sequence. This is the primary entry point
'for AnalysisOutput when building analysis tables.
'@export
Public Sub Build(): End Sub

'@jump:format
'@sub-title Apply formatting to the table using design specifications
'@details
'Applies visual formatting (borders, fills, fonts, number formats) to
'the cross-table cells using the provided ILLFormat design specification.
'Handles header rows, data cells, total rows, and percentage columns
'according to the table type.
'@param designFormat ILLFormat. The formatting specification to apply.
'@export
Public Sub Format(ByVal designFormat As ILLFormat): End Sub
