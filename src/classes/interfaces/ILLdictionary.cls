VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Dictionary Class"

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Dictionary Class")

Option Explicit

'Exposed methods

'The dictionary is based on the datasheet class, so you can extract the datasheet with
'the data Property
Public Property Get Data() As IDataSheet: End Property

'Get the range of the one column of the dictionary
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'A prepared dictionary is a dictionary where the geo lines have been added
'for the geo variables, and some supplementary columns have bee created (table names, etc.)
'This dictionary is a bit differrent from the dictionary in the setup because some stuffs has
'been added. The preparation is mandatory before creating the linelist, an unprepared dictionary
'should not go throught the linelist creation process.

'Know if a dictionary is prepared
Public Property Get Prepared() As Boolean: End Property

'Prepare the dictionary
Public Sub Prepare(ByVal PreservedSheetNames As BetterArray, _ 
                   ByVal geoObject As ILLGeo, _ 
                   ByVal msgTrads As ITranslation, _ 
                   Optional ByVal tablePrefix As String = "table")
End Sub


'Functions
'Check if a column is present in the dictionary
Public Function ColumnExists(ByVal colName As String, _ 
                             Optional ByVal checkValidity As Boolean) As Boolean
End Function

'Return unique values of a column
Public Function UniqueValues(ByVal colName As String) As BetterArray: End Function

'Check if a variable exists in the dictionary
Public Function VariableExists(ByVal varName As String) As Boolean: End Function

'Void Cleaning methods

'Clean the dictionary (use with care)
Public Sub Clean(): End Sub

'Import a datasheet in the dictionary
Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
End Sub

'Export the dictionary to another workbook
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal exportType As String = "__all__",  _ 
                  Optional Byval addListObject As Boolean = True, _
                  Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'Translate the dictionary using a translation object
Public Sub Translate(ByVal TransObject As ITranslation)
End Sub

'Add new column to the dictionary (use with care, the dictionary should normally not be modified)
Public Sub AddColumn(ByVal colName As String): End Sub