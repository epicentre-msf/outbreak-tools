VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of DataSheet Object"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation
'@Interface
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of DataSheet Object")

Option Explicit

'@jump:Create

'@jump:elements
'@section DataSheet Elements
'===============================================================================

'@jump:wksh
Public Property Get Wksh() As Worksheet: End Property

'@jump:startrow
Public Property Get DataStartRow() As Long: End Property

'@jump:startcol
Public Property Get DataStartColumn() As Long: End Property

'@jump:name
Public Property Get Name() As String: End Property

'@jump:endrow
Public Property Get DataEndRow() As Long: End Property

'@jump:endcol
Public Property Get DataEndColumn() As Long: End Property

'@jump:headerrange
Public Property Get HeaderRange() As Range
End Property

'@jump:columnexists
Public Function ColumnExists(ByVal colName As String, _ 
                             Optional ByVal strictSearch As Boolean = True, _ 
                             Optional ByVal matchCase As Boolean = True) As Boolean
End Function

'@jump:columnindex
Public Function ColumnIndex(ByVal colName As String, _
                            Optional ByVal inDataRange As Boolean = False, _
                            Optional shouldExist As Boolean = False, _ 
                            Optional ByVal strictSearch As Boolean = True, _ 
                            Optional ByVal matchCase As Boolean = True) As Long
End Function



'@jump:datarange
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = True, _ 
                              Optional ByVal matchCase As Boolean = True) As Range
End Property

'@jump:RenameColumn
Public Sub RenameColumn(ByVal currentName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)
End Sub

'@jump:operate
'@section Operate with Outside world
'===============================================================================

'@jump:clean
Public Sub Clean() : End Sub

Public Sub AddFormatsColumns(ByVal matchColumnsCase As Boolean, ByVal resetColumns As Boolean, ParamArray lsCols() As Variant)
End Sub

'@jump:filterdata
Public Function FilterData(ByVal varName As String, ByVal criteriaName As String, _
                           ByVal returnedColumnName As String, _
                           Optional ByVal includeHeaders As Boolean = False, _
                           Optional ByVal matchCase As Boolean = True) As BetterArray
End Function

'@jump:filtersdata
'multiple filters
Public Function FiltersData(ByVal varData As BetterArray, _
                            ByVal criteriaData As BetterArray, _
                            ByVal returnedColumnsData As BetterArray) As BetterArray
End Function

'@jump:import
Public Sub Import(ByVal importedData As Object, Optional ByVal strictColumnSearch As Boolean = True): End Sub

'@jump:importformat
Public Sub ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
End Sub

'@jump:export
'Import and Export Sub
'@param toWkb Destination workbook.
'@param filteredVarName Optional column used for filtering prior to export.
'@param filteredCondition Optional condition applied to the filter column.
'@param Hide Optional worksheet visibility after export.
'@param includeNames Optional flag requesting hidden names to be exported alongside worksheet data.
Public Sub Export(ByVal toWkb As Workbook, _
                  Optional ByVal filteredVarName As String = "__all__", _
                  Optional ByVal filteredCondition As String = "<>", _ 
                  Optional ByVal Hide As Long = xlSheetHidden, _
                  Optional ByVal includeNames As Boolean = False)
End Sub


'@section Checkings
'===============================================================================

'@jump:HasCheckings
'Check if the current choice has some checkings
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Return checking values
Public Property Get CheckingValues() As Object: End Property
