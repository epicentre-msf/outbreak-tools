VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for linelist data import operations"

'@Interface
'@Folder("DataIO")
'@ModuleDescription("Interface for linelist data import operations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

Option Explicit

'@class ILLImporter
'@description
'Defines the public contract for the LLImporter class, exposing methods
'to import linelist data from external workbooks. Supports importing
'HList/VList data, custom dropdowns, show/hide state, editable labels,
'single values, and geobase data. Also provides an import report that
'tracks which sheets and variables were not imported or not touched.


'@section Enumerations
'===============================================================================
'@description Scopes for import report entries, used to query which items were missed.

Public Enum ImportReportScope
    ImportReportNotImported = 1
    ImportReportNotTouched = 2
End Enum


'@section Data Checks
'===============================================================================

'@jump:has-data
'@prop-title Whether the workbook contains user-entered data
'@details
'Scans all HList sheets for non-empty rows beyond the blank count
'threshold stored in cell(1,6). Returns True as soon as a single HList
'row has user data. Used to warn the user before clearing or overwriting.
'@return Boolean. True when at least one HList row has user data.
'@export
Public Property Get HasData() As Boolean
End Property

'@jump:has-same-language
'@sub-title Check whether the import workbook language matches
'@details
'Reads the language tag from the import workbook Metadata sheet and
'compares it with the source workbook dictionary language. Stores the
'result internally so that ImportShowHide, ImportEditableLabels, and
'ImportSingleValues silently skip when languages differ.
'@param impwb Workbook. The opened import workbook.
'@return Boolean. True when languages match.
'@export
Public Function HasSameLanguage(ByVal impwb As Workbook) As Boolean
End Function


'@section Data Operations
'===============================================================================

'@jump:clear-data
'@sub-title Clear all user-entered data from HList and VList sheets
'@details
'Iterates over all worksheets in the source workbook, clearing non-formula
'cell values from HList ListObjects and VList named ranges. Unprotects
'and re-protects each sheet via IPasswords.
'@export
Public Sub ClearData()
End Sub

'@jump:import-data
'@sub-title Import data from HList and VList worksheets
'@details
'Iterates over the import workbook worksheets, matching them to source
'workbook sheets by name. For HList sheets, delegates to CustomTable.Import.
'For VList sheets, copies named range values. Tracks imported sheets and
'variables, and builds "not imported" report entries for unmatched items.
'@param impwb Workbook. The opened import workbook.
'@param pasteAtBottom Boolean. When True, appends data below existing rows.
'@export
Public Sub ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
End Sub

'@jump:import-custom-dropdown
'@sub-title Import custom dropdown choices from the import workbook
'@details
'Reads custom dropdown ListObjects from the source workbook custom dropdown
'sheet, matches them to choices in the import workbook Choices sheet, and
'updates each dropdown with the imported values. Silently exits when the
'import workbook has no Choices sheet or no custom dropdown sheet exists.
'@param impwb Workbook. The opened import workbook.
'@param pasteAtBottom Boolean. When True, appends choices below existing entries.
'@export
Public Sub ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
End Sub


'@section Migration Metadata
'===============================================================================

'@jump:import-show-hide
'@sub-title Import show/hide visibility state from a migration workbook
'@details
'Reads the show_hide_migration ListObject from the import workbook and
'applies visibility and column/row sizes to each data sheet via
'ShowHideManager. Silently skips when HasSameLanguage returned False, or
'when the import workbook lacks the required Metadata or show_hide__ sheets,
'or when the metadata tag is not set to "update on import".
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportShowHide(ByVal impwb As Workbook)
End Sub

'@jump:import-editable-labels
'@sub-title Import editable label values from a migration workbook
'@details
'Reads the import workbook dictionary for editable label values and updates
'both the worksheet cells and dictionary entries in the source workbook.
'Silently skips when HasSameLanguage returned False, the import dictionary
'is not prepared, or the metadata tag is not set to "update on import".
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportEditableLabels(ByVal impwb As Workbook)
End Sub

'@jump:import-single-values
'@sub-title Import single metadata values from a migration workbook
'@details
'Reads epiweek start and other single values from the import workbook
'Metadata sheet and updates the corresponding named ranges in the source
'workbook. Silently skips when HasSameLanguage returned False.
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportSingleValues(ByVal impwb As Workbook)
End Sub

'@jump:import-geobase
'@sub-title Import geobase data from a workbook
'@details
'Delegates to LLGeo.Import to load geobase data. When not histoOnly,
'also updates geobase headers in HList worksheets, dictionary geo labels,
'and show/hide tables for geo variables.
'@param impwb Workbook. The opened geobase workbook.
'@param histoOnly Optional Boolean. When True, imports only historic geobase data. Defaults to False.
'@export
Public Sub ImportGeobase(ByVal impwb As Workbook, Optional ByVal histoOnly As Boolean = False)
End Sub


'@section Import Report
'===============================================================================

'@jump:need-report
'@prop-title Whether the last import had issues
'@details
'Returns True when at least one report entry exists, meaning some sheets
'or variables were not imported or not touched during the last import.
'@return Boolean. True when the import report has entries.
'@export
Public Property Get NeedReport() As Boolean
End Property

'@jump:finalize-report
'@sub-title Build the "not touched" part of the import report
'@details
'Compares source workbook sheets and variables against imported items to
'find which ones were not touched by the import. Skips formula-type
'variables (case_when, formula, choice_formula). Call after ImportData.
'@export
Public Sub FinalizeReport()
End Sub

'@jump:report-sheets
'@sub-title Report sheet names for the given scope
'@details
'Returns a BetterArray of sheet names that were either not imported
'(ImportReportNotImported) or not touched (ImportReportNotTouched).
'Returns an empty BetterArray when no entries exist for the scope.
'@param scope Byte. ImportReportNotImported or ImportReportNotTouched.
'@return BetterArray. Sheet names matching the scope, or empty.
'@export
Public Function ReportSheets(ByVal scope As Byte) As BetterArray
End Function

'@jump:report-variables
'@sub-title Report variable entries for the given scope
'@details
'Returns a BetterArray where each entry is a Variant array of
'Array(variableName, sheetName). Returns an empty BetterArray when no
'entries exist for the scope.
'@param scope Byte. ImportReportNotImported or ImportReportNotTouched.
'@return BetterArray. Variant arrays of (variableName, sheetName), or empty.
'@export
Public Function ReportVariables(ByVal scope As Byte) As BetterArray
End Function

'@jump:clear-report
'@sub-title Reset report state
'@details
'Clears all report tracking arrays and resets the needReport flag. Call
'before starting a new import session to ensure a clean report.
'@export
Public Sub ClearReport()
End Sub