VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of exports"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of exports")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument

'@depends Checking

Option Explicit

'@section Properties
'===============================================================================
'@jump:Data
'@prop-title Datasheet abstraction of the export table.
Public Property Get Data() As IDataSheet
End Property

'@jump:Worksheet
'@prop-title Worksheet hosting the export table.
Public Property Get Worksheet() As Worksheet
End Property

'@jump:NumberOfExports
'@prop-title Count of configured export rows.
Public Property Get NumberOfExports() As Long
End Property

'@jump:HasCheckings
'@prop-title Whether informational checkings are available.
Public Property Get HasCheckings() As Boolean
End Property

'@jump:CheckingValues
'@prop-title Collected informational diagnostics about export generation.
Public Property Get CheckingValues() As IChecking
End Property


'@section Rows management
'===============================================================================
'@jump:AddRows
'@prop-title Append export definition rows.
'@param dict Optional dictionary to keep export columns aligned.
Public Sub AddRows(Optional ByVal dict As ILLdictionary)
End Sub

'@jump:InsertRows
'@prop-title Insert export rows based on a target selection.
'@param targetCell Range representing the rows to mirror.
'@param dict Optional dictionary to keep export columns aligned.
'@param insertShift Optional Boolean to insert worksheet rows when needed.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal dict As ILLdictionary, _
                      Optional ByVal insertShift As Boolean = False)
End Sub

'@jump:RemoveRows
'@prop-title Remove trailing export rows.
'@param rowCount Optional Long. Rows with non-empty cell counts at or below this value are removed.
Public Sub RemoveRows(Optional ByVal dict As ILLdictionary, Optional ByVal rowCount As Long = 2)
End Sub

'@jump:ManageRows
Public Sub ManageRows(Optional ByVal del As Boolean = False, _ 
                      Optional ByVal dict As ILLdictionary, _ 
                      Optional ByVal rowCount As Long = 2)
End Sub

'@jump:SyncDictionaryExports
'@prop-title Synchronise dictionary export columns with current export definitions.
'@param dict Optional dictionary reference to align. When omitted, the last cached dictionary is reused.
Public Sub SyncDictionaryExports(Optional ByVal dict As ILLdictionary)
End Sub

'@section SpecManagement
'===============================================================================

'@jump:ImportSpecs
'@prop-title Import export specifications from another worksheet.
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, _
                       ByVal fromStartRow As Long, _
                       ByVal fromStartColumn As Long)
End Sub

'@jump:ExportSpecs
'@prop-title Export specifications into another workbook.
Public Sub ExportSpecs(ByVal toWkb As Workbook, _
                       Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'@jump:InvalidateCaches
'@prop-title Reset cached column lookups and diagnostics.
Public Sub InvalidateCaches()
End Sub

'@section FileName
'===============================================================================
'@jump:ExportFileName
'@prop-title Build a filename for one export definition.
'@param exportNumber Export definition identifier.
'@param dict Dictionary supplying metadata for chunk substitution.
'@param pass Password provider for version metadata.
'@param exportAll Optional flag to force the export-all naming scheme.
Public Function ExportFileName(ByVal exportNumber As Long, _
                               ByVal dict As ILLdictionary, _
                               ByVal pass As IPasswords, _
                               Optional ByVal exportAll As Boolean = False) As String
End Function

'@section Utilities
'===============================================================================
'@jump:IsActive
'@prop-title Determine if an export definition is marked active.
Public Function IsActive(ByVal exportNumber As Long) As Boolean
End Function

'@jump:ColumnValue
'@prop-title Retrieve a column value for a given export row.
Public Function ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
End Function

'@jump:ActiveExportNumbers
'@prop-title Enumerate export numbers that are currently active.
Public Function ActiveExportNumbers() As BetterArray
End Function

'@jump:Translate
Public Sub Translate(ByVal trads As ITranslationObject): End Sub

'@jump:ResetCaches
Public Sub ResetCaches(): End Sub
