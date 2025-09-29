VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the unified linelist export handler"

'@Interface
'@Folder("Exports")
'@ModuleDescription("Interface describing the unified linelist export handler")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument


'@depends Checking

Option Explicit

'@section Factory
'===============================================================================
'@label:Create
'@prop-title Create a new export handler instance
'@param exportSheet Worksheet hosting the export specification table.
'@param startRow Optional start row of the specification table (defaults to 1).
'@param startColumn Optional start column of the specification table (defaults to 1).
'@return An instance implementing LL export coordination.
Public Function Create(ByVal exportSheet As Worksheet, _
                       Optional ByVal startRow As Long = 1, _
                       Optional ByVal startColumn As Long = 1) As ILLExport
End Function

'@section Properties
'===============================================================================
'@label:Data
'@prop-title Datasheet abstraction of the export table.
Public Property Get Data() As IDataSheet
End Property

'@label:Worksheet
'@prop-title Worksheet hosting the export table.
Public Property Get Worksheet() As Worksheet
End Property

'@label:NumberOfExports
'@prop-title Count of configured export rows.
Public Property Get NumberOfExports() As Long
End Property

'@label:HasCheckings
'@prop-title Whether informational checkings are available.
Public Property Get HasCheckings() As Boolean
End Property

'@label:CheckingValues
'@prop-title Collected informational diagnostics about export generation.
Public Property Get CheckingValues() As IChecking
End Property

'@section SpecManagement
'===============================================================================
'@label:AddRows
'@prop-title Append export definition rows.
'@param rowCount Optional number of rows to add (defaults to 1).
Public Sub AddRows(Optional ByVal rowCount As Long = 1)
End Sub

'@label:RemoveRows
'@prop-title Remove trailing export rows.
'@param totalCount Optional number of rows to keep; pass 0 to trim empty rows.
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0)
End Sub

'@label:ImportSpecs
'@prop-title Import export specifications from another worksheet.
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, _
                       ByVal fromStartRow As Long, _
                       ByVal fromStartColumn As Long)
End Sub

'@label:ExportSpecs
'@prop-title Export specifications into another workbook.
Public Sub ExportSpecs(ByVal toWkb As Workbook, _
                       Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'@label:InvalidateCaches
'@prop-title Reset cached column lookups and diagnostics.
Public Sub InvalidateCaches()
End Sub

'@section FileName
'===============================================================================
'@label:ExportFileName
'@prop-title Build a filename for one export definition.
'@param exportNumber Export definition identifier.
'@param dict Dictionary supplying metadata for chunk substitution.
'@param pass Password provider for version metadata.
'@param exportAll Optional flag to force the export-all naming scheme.
Public Function ExportFileName(ByVal exportNumber As Long, _
                               ByVal dict As ILLdictionary, _
                               ByVal pass As ILLPasswords, _
                               Optional ByVal exportAll As Boolean = False) As String
End Function

'@section Utilities
'===============================================================================
'@label:IsActive
'@prop-title Determine if an export definition is marked active.
Public Function IsActive(ByVal exportNumber As Long) As Boolean
End Function

'@label:ColumnValue
'@prop-title Retrieve a column value for a given export row.
Public Function ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
End Function

'@label:ActiveExportNumbers
'@prop-title Enumerate export numbers that are currently active.
Public Function ActiveExportNumbers() As BetterArray
End Function
