VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract describing the linelist workbook facade"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract describing the linelist workbook facade")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILinelist
'@description
'Defines the public contract for the Linelist class, exposing worksheet
'accessors, collaborator getters, and lifecycle methods for creating,
'populating, saving, and error-managing the output linelist workbook.


'@section Worksheets
'===============================================================================
'@description Worksheet resolution by name with optional scope prefix.

'@jump:wksh
'@prop-title Worksheet from the output workbook by name
'@details
'Returns the worksheet matching the given name from the output workbook.
'Raises an error if the sheet does not exist.
'@param sheetName String. The worksheet name to look up.
'@return Worksheet. The resolved worksheet.
Public Property Get Wksh(ByVal sheetName As String) As Worksheet
End Property

'@jump:print-wksh
'@prop-title Printed companion worksheet by base name
'@details
'Returns the printed companion worksheet (prefixed with "print_") from the
'output workbook. Used for HList sheets that have a printable duplicate.
'@param sheetName String. The base sheet name (without the print_ prefix).
'@return Worksheet. The printed companion worksheet.
Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
End Property

'@jump:crf-wksh
'@prop-title CRF companion worksheet by base name
'@details
'Returns the CRF companion worksheet (prefixed with "crf_") from the output
'workbook. Used for case report form layouts.
'@param sheetName String. The base sheet name (without the crf_ prefix).
'@return Worksheet. The CRF companion worksheet.
Public Property Get CRFWksh(ByVal sheetName As String) As Worksheet
End Property

'@jump:sheet-exists
'@sub-title Check whether a sheet exists in the output workbook
'@details
'Tests whether a worksheet with the given name exists in the output workbook
'without raising an error. Uses On Error Resume Next internally.
'@param sheetName String. The worksheet name to test.
'@return Boolean. True when the sheet exists, False otherwise.
Public Function SheetExists(ByVal sheetName As String) As Boolean
End Function

'@jump:add-output-sheet
'@sub-title Add a worksheet to the output workbook
'@details
'Creates a new worksheet in the output workbook with the given name, applying
'an optional scope prefix ("print_" or "crf_") and visibility setting. The
'sheet name is truncated to 30 characters. Raises an error if a sheet with
'the resolved name already exists.
'@param sheetName String. The base name for the new worksheet.
'@param Hide Optional Long. Excel visibility constant. Defaults to xlSheetVeryHidden.
'@param sheetScope Optional Byte. 1 = no prefix, 2 = print_ prefix, 3 = crf_ prefix.
'Defaults to 1.
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1)
End Sub


'@section Collaborators
'===============================================================================
'@description Accessors for helper objects managed by the linelist.

'@jump:dropdown
'@prop-title Dropdown lists manager for the output workbook
'@details
'Returns an IDropdownLists instance backed by the appropriate list sheet.
'Scope 1 uses the temporary dropdown_lists__ sheet with "dropdown_" prefix;
'scope 2 uses the custom choice sheet with no prefix.
'@param scope Optional Byte. 1 = standard dropdowns, 2 = custom choice dropdowns.
'Defaults to 1.
'@return IDropdownLists. The dropdown lists manager.
Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
End Property

'@jump:pivots
'@prop-title Pivot tables manager for the output workbook
'@details
'Returns an ICustomPivotTable instance backed by the translated custom pivot
'table sheet in the output workbook.
'@return ICustomPivotTable. The pivot tables manager.
Public Property Get Pivots() As ICustomPivotTable
End Property

'@jump:linelist-data
'@prop-title Specifications backing this linelist
'@details
'Returns the ILinelistSpecs instance that was supplied during factory
'construction, providing access to dictionary, choices, translations, and
'all other setup data.
'@return ILinelistSpecs. The linelist specifications.
Public Property Get LinelistData() As ILinelistSpecs
End Property

'@jump:dictionary
'@prop-title Dictionary from the specifications
'@details
'Convenience accessor that delegates to LinelistData.Dictionary.
'@return ILLdictionary. The linelist dictionary.
Public Property Get Dictionary() As ILLdictionary
End Property


'@section Lifecycle
'===============================================================================
'@description Build, save, and error-manage the output workbook.

'@jump:prepare
'@sub-title Build the full linelist output workbook from specifications
'@details
'Orchestrates the complete linelist creation process: creates the output
'workbook (or opens a template), adds temporary sheets, prepares hidden names
'and ListObjects, exports specifications, adds analysis and data entry sheets,
'prepares the admin sheet, and transfers all VBA code components.
Public Sub Prepare()
End Sub

'@jump:save-ll
'@sub-title Save the output workbook as .xlsb with password protection
'@details
'Saves the output workbook in .xlsb format with the configured password,
'protects the workbook structure, activates an appropriate visible sheet,
'closes the workbook, and cleans up temporary files.
Public Sub SaveLL()
End Sub

'@jump:error-manage
'@sub-title Display an error dialog for incomplete linelist operations
'@details
'Shows a message box with the error details and asks the user whether they
'want to view the incomplete linelist workbook or close it without saving.
'@param textMessage Optional String. Error description to display.
'Defaults to vbNullString.
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub
