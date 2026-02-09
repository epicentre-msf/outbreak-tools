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


'@section Enumerations
'===============================================================================

' @description Scopes for export operations.
Public Enum ExportScope
    ExportScopeCustom = 1
    ExportScopeMigration = 2
    ExportScopeAnalysis = 3
End Enum


'@section Export methods
'===============================================================================

' @description Export a custom export (export 1, 2, 3, etc.).
' Creates an output workbook with filtered data and metadata, saves to folderPath.
' @param exportNumber Long. Which export definition to use (1-based).
' @param folderPath String. Target folder for the saved file.
' @param useFilters Boolean. When True, exports filtered HList data instead of raw data.
' @return String. Saved file path; empty string on failure.
Public Function ExportCustom(ByVal exportNumber As Long, _
                             ByVal folderPath As String, _
                             Optional ByVal useFilters As Boolean = False) As String
End Function

' @description Export all data for migration to another linelist.
' Creates an output workbook with full data and metadata tags for re-import.
' @param folderPath String. Target folder for the saved file.
' @param includeShowHide Boolean. When True, includes show/hide state with column widths.
' @param keepLabels Boolean. When True, marks editable labels for update on import.
' @return String. Saved file path; empty string on failure.
Public Function ExportMigration(ByVal folderPath As String, _
                                Optional ByVal includeShowHide As Boolean = False, _
                                Optional ByVal keepLabels As Boolean = False) As String
End Function

' @description Export analysis worksheets only.
' @param folderPath String. Target folder for the saved file.
' @return String. Saved file path; empty string on failure.
Public Function ExportAnalysis(ByVal folderPath As String) As String
End Function

' @description Export geobase data to a separate workbook.
' @param folderPath String. Target folder for the saved file.
' @param onlyHistoric Boolean. When True, exports only the historic geobase.
' @return String. Saved file path; empty string on failure.
Public Function ExportGeo(ByVal folderPath As String, _
                           Optional ByVal onlyHistoric As Boolean = False) As String
End Function


'@section State
'===============================================================================

' @description Password used in the last export operation (for display to the user).
' Empty string when no password was used.
Public Property Get LastExportPassword() As String
End Property


'@section Cleanup
'===============================================================================

' @description Close all output workbooks that may still be open after an error.
Public Sub CloseAll()
End Sub
