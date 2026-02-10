VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomPivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for custom pivot table creation on a dedicated worksheet"

'@Interface
'@Folder("Tables")
'@ModuleDescription("Interface for custom pivot table creation on a dedicated worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ICustomPivotTable
'@description
'Defines the public contract for the CustomPivotTable class.
'Manages pivot table creation on a dedicated worksheet during linelist
'building. Each call to Add creates a new PivotTable block (title + pivot)
'stacked vertically on the sheet.


'@section Methods
'===============================================================================
'@description Pivot table creation and formatting.

'@jump:add
'@sub-title Create a pivot table from a ListObject source and add a titled block
'@details
'Creates a PivotTable using PivotTableWizard with the given table as data
'source. Writes a styled title above the pivot and stores the title cell
'address as a worksheet-level HiddenName (RNG_PivotTitle_<tableName>).
'Successive calls stack blocks vertically on the sheet.
'@param title String. Display label for the pivot block (typically the sheet name).
'@param tableName String. Name of the source ListObject providing pivot data.
'@param pivotName String. Localised prefix for the title (e.g. "Pivot Table"). Defaults to "Pivot Table".
Public Sub Add(ByVal title As String, _
               ByVal tableName As String, _
               Optional ByVal pivotName As String = "Pivot Table")
End Sub

'@jump:format
'@sub-title Apply design formatting to the pivot table worksheet
'@details
'Delegates to ILLFormat.ApplyFormat with the HListCustomPivotTables scope
'to style the entire pivot worksheet according to the linelist design.
'@param design ILLFormat. The format object providing style rules.
Public Sub Format(ByVal design As ILLFormat)
End Sub
