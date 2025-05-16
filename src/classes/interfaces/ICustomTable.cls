VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of CustomTable"

'@Folder("GeneralInterface")
'@ModuleDescription("Interface of CustomTable")
'@Interface
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

'@jump:Create

Option Explicit

'@section CustomTable elements

'@jump:HeaderRange
'Header row range
Public Property Get HeaderRange() As Range: End Property

'@jump:DataRange
'Range of a column
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = False) _
                              As Range
End Property

'@jump:Name
'Name of the attached listObject
Public Property Get Name() As String: End Property

'@jump:CellRange
'Specific cell in the customtable
Public Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range: End Property

'Specific value of a customtable with ID column
Public Property Get Value(ByVal colName As String, ByVal keyName As String) As String: End Property

'@section Modify the table

'@jump:AddIds
'Add Ids to the ID Column in a sequential way
Public Sub AddIds(): End Sub

'@jump:AddRows
'Add Rows to the custom table
Public Sub AddRows(Optional ByVal nbRows As Long = 5, _ 
                   Optional ByVal insertShift As Boolean = False)
End Sub

'@jump:RemoveRows
'Remove rows to the custom table
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0, Optional ByVal includeIds As Boolean = True)
End Sub

'@jump:SetValidation
'Set validation on one specific column of the listObject
Public Sub SetValidation(ByVal colName As String, _
                         ByVal drop As IDropdownLists, _
                         ByVal dropName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString)
End Sub

'@jump:SetValue
'Use very wisely, could modify the custom table values
Public Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)
End Sub

'@jump:sort
'Sort columns
Public Sub Sort(Optional ByVal colName As String = vbNullString, _ 
                Optional ByVal colList As Object = Nothing, _ 
                Optional ByVal directSort As Boolean = True)
End Sub

'@jump:Clean
Public Sub Clean(): End Sub

'@section Import Export

'@jump:Import
'ImpTab could be either a customTable or a DataSheet Object
'pasteAtBottom: paste the new imported data at the bottom of the previous one
'strictColumnSearch: When looking for columns matching in import, do it the strict way (case sensitive)
'insertShift: When resizing before import, insert some shift (maybe time consuming) instead of extending the
'table.
Public Sub Import(ByVal impTab As Object, _
                  Optional ByVal pasteAtBottom As Boolean = False, _
                  Optional ByVal strictColumnSearch As Boolean = False, _ 
                  Optional ByVal insertShift As Boolean = True, _
                  Optional ByVal keepImportHeaders As Boolean = False)
End Sub


'@jump:ImportAll
Public Sub ImportAll(ByVal impTab As Object): End Sub


'@jump:ImportColumnsNotFound
Public Property Get ImportColumnsNotFound() As BetterArray
End Property

'@jump:HasColumnsNotImported
'Test if there are imports in the custom table
Public Property Get HasColumnsNotImported() As Boolean
End Property

'@jump:Export
Public Sub Export(ByVal sh As Worksheet,  _ 
                  Optional ByVal headersTable As Object = Nothing, _ 
                  Optional ByVal startLine As Long = 1)
End Sub

'@section Checkings

'@jump:HasCheckings
'Check if the current choice has some checkings
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Return checking values
Public Property Get CheckingValues() As Object: End Property