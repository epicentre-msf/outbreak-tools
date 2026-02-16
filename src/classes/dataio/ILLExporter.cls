VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for linelist data export operations"

'@Interface
'@Folder("DataIO")
'@ModuleDescription("Interface for linelist data export operations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

Option Explicit

'@class ILLExporter
'@description
'Defines the public contract for the LLExporter class, exposing methods
'to export linelist data into standalone workbooks. Supports custom
'filtered exports, full migration exports, analysis-only exports, and
'geobase exports. Consumers use ExportCustom, ExportMigration,
'ExportAnalysis, or ExportGeo, then read LastExportPassword if a
'password was applied.


'@section Enumerations
'===============================================================================
'@description Scopes for export operations, used to tag metadata in output workbooks.

Public Enum ExportScope
    ExportScopeCustom = 1
    ExportScopeMigration = 2
    ExportScopeAnalysis = 3
End Enum


'@section Export Methods
'===============================================================================

'@jump:export-custom
'@sub-title Export a numbered custom export definition
'@details
'Creates an output workbook containing filtered data, optional metadata
'sheets, and optional analysis sheets according to the export definition
'identified by exportNumber. Saves the workbook to folderPath and returns
'the full file path. When useFilters is True, exports filtered HList data
'from companion filter sheets rather than the raw ListObject data.
'@param exportNumber Long. Which export definition to use (1-based).
'@param folderPath String. Target folder for the saved file.
'@param useFilters Optional Boolean. When True, exports filtered HList data instead of raw data. Defaults to False.
'@return String. Saved file path, or empty string on failure.
'@export
Public Function ExportCustom(ByVal exportNumber As Long, _
                             ByVal folderPath As String, _
                             Optional ByVal useFilters As Boolean = False) As String
End Function

'@jump:export-migration
'@sub-title Export all data for migration to another linelist
'@details
'Creates an output workbook with full unfiltered data, metadata sheets,
'dictionary, choices, translations, and metadata tags for re-import.
'Optionally includes show/hide state with column widths and marks
'editable labels for update on import. No password is applied.
'@param folderPath String. Target folder for the saved file.
'@param includeShowHide Optional Boolean. When True, includes show/hide state with column widths. Defaults to False.
'@param keepLabels Optional Boolean. When True, marks editable labels for update on import. Defaults to False.
'@return String. Saved file path, or empty string on failure.
'@export
Public Function ExportMigration(ByVal folderPath As String, _
                                Optional ByVal includeShowHide As Boolean = False, _
                                Optional ByVal keepLabels As Boolean = False) As String
End Function

'@jump:export-analysis
'@sub-title Export analysis worksheets only
'@details
'Creates an output workbook containing only the four analysis sheets
'(univariate, time series, spatial, spatio-temporal) with their values,
'formatting, and graphs. No data or metadata sheets are included.
'@param folderPath String. Target folder for the saved file.
'@return String. Saved file path, or empty string on failure.
'@export
Public Function ExportAnalysis(ByVal folderPath As String) As String
End Function

'@jump:export-geo
'@sub-title Export geobase data to a separate workbook
'@details
'Creates an output workbook and delegates to LLGeo.ExportToFile to write
'geobase data. When onlyHistoric is True, exports only the historic
'geobase subset. The file name includes a date stamp.
'@param folderPath String. Target folder for the saved file.
'@param onlyHistoric Optional Boolean. When True, exports only the historic geobase. Defaults to False.
'@return String. Saved file path, or empty string on failure.
'@export
Public Function ExportGeo(ByVal folderPath As String, _
                           Optional ByVal onlyHistoric As Boolean = False) As String
End Function


'@section State
'===============================================================================

'@jump:last-export-password
'@prop-title Password used in the last export operation
'@details
'Returns the password that was applied to the most recent custom export,
'or an empty string when no password was used. Modules display this value
'to the user after a successful export so they can share it securely.
'@return String. The password applied to the last export, or vbNullString.
'@export
Public Property Get LastExportPassword() As String
End Property


'@section Cleanup
'===============================================================================

'@jump:close-all
'@sub-title Close all output workbooks that may still be open
'@details
'Closes any output workbook left open after an export error, preventing
'orphaned minimized windows. Safe to call even when no workbook is open.
'@export
Public Sub CloseAll()
End Sub