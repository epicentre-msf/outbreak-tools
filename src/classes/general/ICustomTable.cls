VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of CustomTable"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of CustomTable")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ICustomTable
'@description
'Defines the public contract for the CustomTable class, exposing
'read/write helpers around an Excel ListObject including column and
'cell access, row manipulation, sorting, validation, import/export,
'and diagnostic logging.


'@section Elements
'===============================================================================
'@description Properties that expose the table structure and cell access.

'@jump:header-range
'@prop-title Header row Range of the CustomTable
'@details
'Returns the Range covering the header row of the underlying ListObject.
'@return Range. The header row range.
Public Property Get HeaderRange() As Range: End Property

'@jump:id-value
'@prop-title ID column name
'@details
'Returns the name of the column designated as the unique row identifier.
'@return String. The ID column name, or vbNullString when none is set.
Public Property Get IdValue() As String: End Property

'@jump:data-range
'@prop-title Range of a column or the entire data body
'@details
'Returns a Range for the specified column, or the entire data body when
'colName is "__all__". Returns Nothing when the column has no data rows
'or is not found.
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@return Range. The requested data range, or Nothing when unavailable.
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
          Optional ByVal includeHeaders As Boolean = False, _
          Optional ByVal strictSearch As Boolean = False, _
          Optional ByVal matchCase As Boolean = False) As Range
End Property

'@jump:name
'@prop-title Name of the attached ListObject
'@details
'Returns the cached name of the ListObject assigned during creation.
'@return String. The table name.
Public Property Get Name() As String: End Property

'@jump:cell-range
'@prop-title A specific cell in the CustomTable
'@details
'Returns the cell at the intersection of the given column and worksheet
'row number. Returns Nothing when the column does not exist or the row
'is out of bounds.
'@param colName String. Column header identifying the column.
'@param lineNum Long. Worksheet row number (1-based).
'@return Range. The target cell, or Nothing when not found.
Public Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range: End Property

'@jump:value
'@prop-title Retrieve a cell value by column and key
'@details
'For tables with a key column, returns the cell value at the intersection
'of the specified column and the row matching keyName. Returns an empty
'string when the key or column is not found.
'@param colName String. Column header to read from.
'@param keyName String. Key value identifying the row.
'@return String. The cell value converted to string, or vbNullString.
Public Property Get Value(ByVal colName As String, ByVal keyName As String) As String: End Property

'@jump:value-range
'@prop-title Range of a cell identified by column and key
'@details
'Returns the Range at the intersection of the specified column and the
'row matching keyName in the ID column. Returns Nothing when the key or
'column is not found.
'@param colName String. Column header identifying the column.
'@param keyName String. Key value identifying the row.
'@return Range. The target cell range, or Nothing when not found.
Public Property Get ValueRange(ByVal colName As String, ByVal keyName As String) As Range
End Property


'@section Modify
'===============================================================================
'@description Methods that add, remove, or reorganise table rows and columns.

'@jump:add-ids
'@sub-title Add sequential IDs to the ID column
'@details
'Fills the ID column with sequential values, optionally prefixed.
'Silently exits when no ID column is configured.
Public Sub AddIds(): End Sub

'@jump:add-rows
'@sub-title Add rows to the CustomTable
'@details
'Appends the requested number of rows, either by inserting worksheet
'rows (to protect stacked tables) or by extending the ListObject range.
'Optionally renumbers the ID column afterwards.
'@param nbRows Optional Long. Number of rows to add. Defaults to 5.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
Public Sub AddRows(Optional ByVal nbRows As Long = 5, _
                   Optional ByVal insertShift As Boolean = False, _
                   Optional ByVal includeIds As Boolean = True)
End Sub

'@jump:remove-rows
'@sub-title Remove empty rows from the CustomTable
'@details
'Deletes rows whose non-empty cell count falls at or below totalCount.
'Optionally renumbers the ID column afterwards.
'@param totalCount Optional Long. Maximum non-empty cells for a row to be considered empty. Defaults to 0.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = False)
End Sub

'@jump:set-validation
'@sub-title Set dropdown validation on a column
'@details
'Applies data validation from the supplied IDropdownLists to every cell
'in the specified column. Silently exits when the column does not exist.
'@param colName String. Column header to validate.
'@param drop IDropdownLists. Dropdown source providing the validation list.
'@param dropName String. Name of the dropdown list within the source.
'@param alertType Optional String. Alert severity ("info", "warning", "error"). Defaults to "info".
'@param message Optional String. Message shown when validation fails. Defaults to vbNullString.
Public Sub SetValidation(ByVal colName As String, _
                         ByVal drop As IDropdownLists, _
                         ByVal dropName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString)
End Sub

'@jump:set-value
'@sub-title Change the value of one specific cell
'@details
'Writes newValue to the cell identified by column name and key.
'Silently exits when the key or column is not found.
'@param colName String. Column header identifying the cell.
'@param keyName String. Key value identifying the row.
'@param newValue String. The new value to assign.
Public Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)
End Sub

'@jump:rename-column
'@sub-title Rename a column header within the table
'@details
'Replaces the header text of the matched column with newName and
'resets the lookup cache.
'@param colName String. Current column header to find.
'@param newName String. New header text to assign.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
Public Sub RenameColumn(ByVal colName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)
End Sub

'@jump:sort
'@sub-title Sort the table on one or more columns
'@details
'Sorts in ascending order on the specified column, then on each column
'in colList. When directSort is True, uses a simple Range.Sort; when
'False, groups rows by first occurrence of each distinct value.
'@param colName Optional String. Column header to sort on. Defaults to vbNullString.
'@param colList Optional Object. BetterArray of additional column headers. Defaults to Nothing.
'@param directSort Optional Boolean. When True, sorts directly. Defaults to True.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
Public Sub Sort(Optional ByVal colName As String = vbNullString, _
                Optional ByVal colList As Object = Nothing, _
                Optional ByVal directSort As Boolean = True, _
                Optional ByVal strictSearch As Boolean = True)
End Sub

'@jump:clean
'@sub-title Clear all non-formula, non-key columns
'@details
'Clears the contents of every data column except those containing
'formulas or serving as the key column.
Public Sub Clean(): End Sub

'@jump:insert-rows-at
'@sub-title Insert rows at the selected position
'@details
'Inserts rows matching the height of targetCell. Uses worksheet row
'insertion when insertShift is True to protect stacked tables.
'@param targetCell Range. Selection anchoring the insertion point.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
Public Sub InsertRowsAt(ByVal targetCell As Range, _
                        Optional ByVal insertShift As Boolean = False, _
                        Optional ByVal includeIds As Boolean = True)
End Sub

'@jump:delete-rows-at
'@sub-title Delete rows intersecting the selection
'@details
'Removes data rows that intersect targetCell. Uses worksheet row
'deletion when forceShift or the stored shift tracker is active.
'@param targetCell Range. Selection identifying rows to delete.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
Public Sub DeleteRowsAt(ByVal targetCell As Range, _
                        Optional ByVal includeIds As Boolean = True, _
                        Optional ByVal forceShift As Boolean = False)
End Sub


'@section DataExchange
'===============================================================================
'@description Import, export, and snapshot operations.

'@jump:import
'@sub-title Import data from a DataSheet or CustomTable
'@details
'Copies column-matched data from the source into this table. When
'keepSourceHeaders is True, replaces all headers and data entirely.
'Creates a snapshot before import and rolls back on error.
'@param impTab Object. A DataSheet or CustomTable to import from.
'@param pasteAtBottom Optional Boolean. When True, appends below existing data. Defaults to False.
'@param strictColumnSearch Optional Boolean. When True, uses exact column matching. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows during resize. Defaults to True.
'@param formatHeaders Optional Object. BetterArray of headers whose formatting to preserve. Defaults to Nothing.
'@param keepSourceHeaders Optional Boolean. When True, replaces destination headers. Defaults to False.
Public Sub Import(ByVal impTab As Object, _
                  Optional ByVal pasteAtBottom As Boolean = False, _
                  Optional ByVal strictColumnSearch As Boolean = False, _
                  Optional ByVal insertShift As Boolean = True, _
                  Optional ByVal formatHeaders As Object = Nothing, _
                  Optional ByVal keepSourceHeaders As Boolean = False)
End Sub

'@jump:import-all
'@sub-title Replace the entire table with imported data
'@details
'Replaces all headers and data with the content of the supplied object,
'discarding the current structure entirely.
'@param impTab Object. A DataSheet or CustomTable to import from.
Public Sub ImportAll(ByVal impTab As Object): End Sub

'@jump:restore-table-snapshot
'@sub-title Restore the table from a saved snapshot
'@details
'Reverts the table headers and data to the state captured by
'SaveTableSnapshot. Used for rollback after a failed import.
Public Sub RestoreTableSnapshot(): End Sub

'@jump:save-table-snapshot
'@sub-title Save a snapshot of the current table state
'@details
'Captures the current headers and data rows so they can be restored
'later if an import fails.
Public Sub SaveTableSnapshot(): End Sub

'@jump:import-columns-not-found
'@prop-title Columns not matched during the last import
'@details
'Returns a BetterArray of column headers from the source that had no
'match in this table. Only available after a successful import.
'@return BetterArray. Unmatched column headers, or Nothing.
Public Property Get ImportColumnsNotFound() As BetterArray
End Property

'@jump:has-columns-not-imported
'@prop-title Whether the last import had unmatched columns
'@details
'Returns True when the most recent import left some source columns
'unmatched against this table.
'@return Boolean. True when unmatched columns exist.
Public Property Get HasColumnsNotImported() As Boolean
End Property

'@jump:export
'@sub-title Export the table to a worksheet
'@details
'Writes the requested columns to the target worksheet, optionally
'creating a ListObject on the output range.
'@param sh Worksheet. Target worksheet to write to.
'@param headersTable Optional Object. BetterArray of column headers to export. Defaults to Nothing (all columns).
'@param startLine Optional Long. Row to begin writing at. Defaults to 1.
'@param startColumn Optional Long. Column to begin writing at. Defaults to 1.
'@param addListObject Optional Boolean. When True, creates a ListObject on the output. Defaults to False.
Public Sub Export(ByVal sh As Worksheet,  _
                  Optional ByVal headersTable As Object = Nothing, _
                  Optional ByVal startLine As Long = 1, _
                  Optional ByVal startColumn As Long = 1, _
                  Optional ByVal addListObject As Boolean = False, _ 
                  Optional ByVal clearSheet As Boolean = True)
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the table has logged diagnostic messages
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
