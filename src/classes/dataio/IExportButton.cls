VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExportButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for export button wrappers"

'@Interface
'@Folder("DataIO")
'@ModuleDescription("Interface for export button wrappers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

Option Explicit

'@class IExportButton
'@description
'Defines the public contract for the ExportButton class, exposing
'read-only export identification and read/write filter state.
'Consumers query ExportNumber to identify which export definition
'the button drives, and UseFilter to read or toggle the associated
'filter checkbox state.


'@section Properties
'===============================================================================

'@jump:export-number
'@prop-title Export number derived from button name
'@details
'Returns the 1-based export definition number parsed from the
'button Name property (e.g. "CMDExport1" yields 1).
'@return Long. The 1-based export number.
'@export
Public Property Get ExportNumber() As Long
End Property

'@jump:use-filter
'@prop-title Whether the filter checkbox is checked
'@details
'Returns the current value of the filter checkbox. Returns False
'when no checkbox was bound.
'@return Boolean. True when the filter checkbox is checked.
'@export
Public Property Get UseFilter() As Boolean
End Property

'@jump:use-filter-set
'@prop-title Set the filter checkbox value
'@param value Boolean. True to check, False to uncheck.
Public Property Let UseFilter(ByVal value As Boolean)
End Property
