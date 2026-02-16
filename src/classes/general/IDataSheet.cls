VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of DataSheet Object"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of DataSheet Object")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IDataSheet
'@description
'Defines the public contract for the DataSheet class, exposing low-level
'worksheet data access, column lookup, filtering, import/export, and
'diagnostic logging for table-like structures that do not rely on
'ListObjects.


'@section Elements
'===============================================================================
'@description Properties that expose the worksheet layout and column access.

'@jump:wksh
'@prop-title Worksheet backing the DataSheet
'@details
'Returns the worksheet that hosts the table-like data.
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet: End Property

'@jump:data-start-row
'@prop-title First row of data (header row)
'@details
'Returns the 1-based row index where the data starts.
'@return Long. The header row index.
Public Property Get DataStartRow() As Long: End Property

'@jump:data-start-column
'@prop-title First column of data
'@details
'Returns the 1-based column index where the data starts.
'@return Long. The starting column index.
Public Property Get DataStartColumn() As Long: End Property

'@jump:name
'@prop-title Name of the DataSheet
'@details
'Returns the assigned object name, or the worksheet name when none
'was provided during creation.
'@return String. The DataSheet name.
Public Property Get Name() As String: End Property

'@jump:data-end-row
'@prop-title Last row containing data
'@details
'Dynamically computes the last occupied row. When StrictEnd is True,
'scans all non-formula columns for the furthest row.
'@return Long. The 1-based last data row index.
Public Property Get DataEndRow() As Long: End Property

'@jump:data-end-column
'@prop-title Last column containing data
'@details
'Dynamically computes the rightmost occupied column in the header row.
'@return Long. The 1-based last column index.
Public Property Get DataEndColumn() As Long: End Property

'@jump:header-range
'@prop-title Header row Range
'@details
'Returns the Range spanning from the start column to the end column
'on the header row.
'@return Range. The header row range.
Public Property Get HeaderRange() As Range
End Property

'@jump:column-exists
'@sub-title Check whether a column exists in the DataSheet
'@details
'Searches the header row for the specified column name.
'@param colName String. Column header to search for.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Boolean. True when the column is found.
Public Function ColumnExists(ByVal colName As String, _
                             Optional ByVal strictSearch As Boolean = True, _
                             Optional ByVal matchCase As Boolean = True) As Boolean
End Function

'@jump:column-index
'@sub-title Resolve the column index for a header name
'@details
'Returns the worksheet column index (or data-range-relative index when
'inDataRange is True). Returns -1 when the column is not found and
'shouldExist is False; raises an error when shouldExist is True.
'@param colName String. Column header to look up.
'@param inDataRange Optional Boolean. When True, returns the index relative to the data range. Defaults to False.
'@param shouldExist Optional Boolean. When True, raises an error if the column is not found. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Long. The column index, or -1 when not found.
Public Function ColumnIndex(ByVal colName As String, _
                            Optional ByVal inDataRange As Boolean = False, _
                            Optional shouldExist As Boolean = False, _
                            Optional ByVal strictSearch As Boolean = True, _
                            Optional ByVal matchCase As Boolean = True) As Long
End Function

'@jump:data-range
'@prop-title Range of a column or the entire data body
'@details
'Returns a data Range for the specified column, or the entire data body
'when colName is "__all__". When includeHeaders is True, the header row
'is included.
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Range. The requested data range.
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = True, _
                              Optional ByVal matchCase As Boolean = True) As Range
End Property

'@jump:rename-column
'@sub-title Rename a header in the DataSheet
'@details
'Replaces the header text of the matched column with newName.
'@param currentName String. Current column header to find.
'@param newName String. New header text to assign.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
Public Sub RenameColumn(ByVal currentName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)
End Sub


'@section Operations
'===============================================================================
'@description Filtering, cleaning, import, and export operations.

'@jump:clean
'@sub-title Clear all non-formula columns
'@details
'Clears the contents of every data column that does not contain formulas.
Public Sub Clean() : End Sub

'@jump:add-formats-columns
'@sub-title Register columns for formatting import/export
'@details
'Adds the specified column names to the internal formatting list so that
'Import and Export preserve their visual formatting.
'@param matchColumnsCase Boolean. When True, uses case-sensitive column matching.
'@param resetColumns Boolean. When True, clears the existing list before adding.
'@param lsCols ParamArray Variant. Column names to register for formatting.
Public Sub AddFormatsColumns(ByVal matchColumnsCase As Boolean, ByVal resetColumns As Boolean, ParamArray lsCols() As Variant)
End Sub

'@jump:filter-data
'@sub-title Filter data on a single column and return the result
'@details
'Applies an AutoFilter on varName with criteriaName, extracts the
'visible rows from returnedColumnName, and returns them as a
'BetterArray. The filter is removed after the operation.
'@param varName String. Column to filter on.
'@param criteriaName String. Filter criteria to apply.
'@param returnedColumnName String. Column whose values to return.
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param matchCase Optional Boolean. When True, column lookup is case-sensitive. Defaults to True.
'@return BetterArray. The filtered values.
Public Function FilterData(ByVal varName As String, ByVal criteriaName As String, _
                           ByVal returnedColumnName As String, _
                           Optional ByVal includeHeaders As Boolean = False, _
                           Optional ByVal matchCase As Boolean = True) As BetterArray
End Function

'@jump:filters-data
'@sub-title Filter data on multiple columns and return the result
'@details
'Applies multiple AutoFilters (one per entry in varData/criteriaData)
'and returns the visible values from each column in returnedColumnsData
'as a two-dimensional BetterArray.
'@param varData BetterArray. Column names to filter on.
'@param criteriaData BetterArray. Criteria corresponding to each column.
'@param returnedColumnsData BetterArray. Columns whose values to return.
'@return BetterArray. A two-dimensional array of filtered values.
Public Function FiltersData(ByVal varData As BetterArray, _
                            ByVal criteriaData As BetterArray, _
                            ByVal returnedColumnsData As BetterArray) As BetterArray
End Function

'@jump:import
'@sub-title Import data from another DataSheet or CustomTable
'@details
'Cleans this DataSheet, then copies matching columns from the source.
'Preserves formula columns. Applies format import for registered
'formatting columns.
'@param importedData Object. A DataSheet or CustomTable to import from.
'@param strictColumnSearch Optional Boolean. When True, uses case-sensitive column matching. Defaults to True.
Public Sub Import(ByVal importedData As Object, Optional ByVal strictColumnSearch As Boolean = True): End Sub

'@jump:import-format
'@sub-title Import column formatting from a source object
'@details
'Copies visual formatting (colours, font weight, italic) from the
'source for all columns registered via AddFormatsColumns.
'@param importData Object. A DataSheet or CustomTable to copy formatting from.
'@param matchColumnsCase Optional Boolean. When True, uses case-sensitive column matching. Defaults to True.
Public Sub ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
End Sub

'@jump:export
'@sub-title Export the DataSheet to a workbook
'@details
'Creates (or clears) a worksheet in the target workbook with the same
'name as this DataSheet, then writes all data. Optionally filters rows
'before export. The exported sheet is hidden by default.
'@param toWkb Workbook. Destination workbook.
'@param filteredVarName Optional String. Column used for filtering. Defaults to "__all__" (no filter).
'@param filteredCondition Optional String. Filter criteria. Defaults to "<>".
'@param Hide Optional Long. Worksheet visibility after export. Defaults to xlSheetHidden.
'@param includeNames Optional Boolean. When True, exports hidden names alongside data. Defaults to False.
Public Sub Export(ByVal toWkb As Workbook, _
                  Optional ByVal filteredVarName As String = "__all__", _
                  Optional ByVal filteredCondition As String = "<>", _
                  Optional ByVal Hide As Long = xlSheetHidden, _
                  Optional ByVal includeNames As Boolean = False)
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the DataSheet has logged diagnostic messages
'@details
'Returns True when at least one diagnostic entry has been recorded.
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal IChecking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. An IChecking instance, or Nothing.
Public Property Get CheckingValues() As Object: End Property
