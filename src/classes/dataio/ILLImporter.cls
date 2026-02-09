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


'@section Enumerations
'===============================================================================

' @description Scopes for import report entries.
' ImportReportNotImported: items from the import file not found in the current workbook.
' ImportReportNotTouched: items in the current workbook not present in the import file.
Public Enum ImportReportScope
    ImportReportNotImported = 1
    ImportReportNotTouched = 2
End Enum


'@section Data checks
'===============================================================================

' @description Check whether the workbook contains any user-entered data.
' Scans HList sheets for non-empty rows beyond the blank count threshold.
' @return True when at least one HList row has user data.
Public Property Get HasData() As Boolean
End Property

' @description Check whether the import workbook language matches the source workbook.
' Reads the language tag from the import workbook's Metadata sheet and compares
' it with the source workbook's dictionary language.
' Stores the result internally for use by ImportShowHide, ImportEditableLabels,
' and ImportSingleValues (which silently skip when languages differ).
' @param impwb Workbook. The opened import workbook.
' @return True when languages match.
Public Function HasSameLanguage(ByVal impwb As Workbook) As Boolean
End Function


'@section Data operations
'===============================================================================

' @description Clear all user-entered data from HList and VList sheets.
' Preserves formula cells. Unprotects and re-protects each sheet.
Public Sub ClearData()
End Sub

' @description Import data from HList and VList worksheets of the import workbook.
' Tracks imported sheets and variables. Builds "not imported" report entries
' for sheets and variables not found in the current workbook.
' @param impwb Workbook. The opened import workbook.
' @param pasteAtBottom Boolean. When True, appends data below existing rows.
Public Sub ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
End Sub

' @description Import custom dropdown choices from the import workbook.
' Updates custom dropdowns on the source workbook with values from the
' import workbook's Choices sheet.
' @param impwb Workbook. The opened import workbook (must have a Choices sheet).
' @param pasteAtBottom Boolean. When True, appends choices below existing entries.
Public Sub ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
End Sub


'@section Migration metadata
'===============================================================================

' @description Import show/hide visibility state and column widths from a migration workbook.
' Reads the show_hide_migration ListObject and applies visibility and column/row sizes.
' Silently skips when HasSameLanguage returned False, or when the import workbook
' lacks the required Metadata/show_hide__ sheets, or the metadata tag is not set.
' @param impwb Workbook. The opened migration workbook.
Public Sub ImportShowHide(ByVal impwb As Workbook)
End Sub

' @description Import editable label values from a migration workbook.
' Updates both worksheet cells and dictionary entries for editable labels.
' Silently skips when HasSameLanguage returned False.
' @param impwb Workbook. The opened migration workbook.
Public Sub ImportEditableLabels(ByVal impwb As Workbook)
End Sub

' @description Import single values (epiweek start) from a migration workbook.
' Reads the Metadata sheet and updates named ranges in the source workbook.
' Silently skips when HasSameLanguage returned False.
' @param impwb Workbook. The opened migration workbook.
Public Sub ImportSingleValues(ByVal impwb As Workbook)
End Sub

' @description Import geobase data from the given workbook.
' Delegates to LLGeo.Import. When not histoOnly, also updates headers,
' dictionary labels, and show/hide tables.
' @param impwb Workbook. The opened geobase workbook.
' @param histoOnly Boolean. When True, imports only historic geobase data.
Public Sub ImportGeobase(ByVal impwb As Workbook, Optional ByVal histoOnly As Boolean = False)
End Sub


'@section Import report
'===============================================================================

' @description Whether the last import had issues that should be shown in a report.
' Returns True when at least one report entry exists (sheets or variables).
Public Property Get NeedReport() As Boolean
End Property

' @description Build the "not touched" part of the import report.
' Compares current workbook sheets and variables against imported items.
' Skips formula-type variables (case_when, formula, choice_formula).
' Call after ImportData to complete the report.
Public Sub FinalizeReport()
End Sub

' @description Report sheet names for the given scope.
' @param scope Byte. ImportReportNotImported or ImportReportNotTouched.
' @return BetterArray of String sheet names (empty if none).
Public Function ReportSheets(ByVal scope As Byte) As BetterArray
End Function

' @description Report variable entries for the given scope.
' Each entry is a Variant array: Array(variableName, sheetName).
' @param scope Byte. ImportReportNotImported or ImportReportNotTouched.
' @return BetterArray of Variant arrays (empty if none).
Public Function ReportVariables(ByVal scope As Byte) As BetterArray
End Function

' @description Reset report state. Call before starting a new import session.
Public Sub ClearReport()
End Sub
