VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Dictionary class"
Option Explicit

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Dictionary class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation


'@section Data Access
'===============================================================================

'@jump:data
'@prop-title DataSheet backing store
'@return An IDataSheet reference with the dictionary data
Public Property Get Data() As IDataSheet: End Property

'@jump:datarange
'@prop-title Data range accessor
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. Include headers in the returned range. Defaults to False.
'@return A Range scoped to the requested dictionary data
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'@jump:prepared
'@prop-title Dictionary preparation status
'@return True when the dictionary is ready for designer workflows
Public Property Get Prepared() As Boolean: End Property

'@jump:exports
'@prop-title Total number of export columns
'@return A Long representing the configured number of export slots
Public Property Get TotalNumberOfExports() As Long: End Property


'@section Column Operations
'===============================================================================

'@jump:columnexists
'@prop-title Test for column existence
'@param colName String. Column header to search for.
'@param checkValidity Optional Boolean. When True, the column is validated against the allowed list.
'@return True when the column is present
Public Function ColumnExists(ByVal colName As String, _
                             Optional ByVal checkValidity As Boolean) As Boolean
End Function

'@jump:uniquevalues
'@prop-title Unique column values
'@param colName String. Column header for the lookup.
'@return A BetterArray containing distinct values
Public Function UniqueValues(ByVal colName As String) As BetterArray: End Function

'@jump:variableexists
'@prop-title Variable presence test
'@param varName String. Variable identifier to search for.
'@return True when the variable exists in the dictionary
Public Function VariableExists(ByVal varName As String) As Boolean: End Function

'@jump:addcolumn
'@sub-title Append a column to the dictionary
'@param colName String. Name of the column to append.
Public Sub AddColumn(ByVal colName As String): End Sub

'@jump:insertcolumn
'@sub-title Insert a column relative to another
'@param colName String. Name of the column to insert.
'@param After String. Existing column after which the new column is inserted.
Public Sub InsertColumn(ByVal colName As String, ByVal After As String): End Sub

'@jump:removecolumn
'@sub-title Remove a column from the dictionary
'@param colName String. Column header to delete.
Public Sub RemoveColumn(ByVal colName As String): End Sub


'@section Row Operations
'===============================================================================

'@jump:addrows
'@sub-title Add rows to the dictionary listobject when available
Public Sub AddRows(): End Sub

'@jump:removerows
'@sub-title Remove rows from the dictionary listobject when available
Public Sub RemoveRows(): End Sub


'@section Specialised Views
'===============================================================================

'@jump:choicesvars
'@prop-title Choice variables list
'@return A BetterArray listing variables controlled by choices
Public Property Get ChoicesVars() As BetterArray: End Property

'@jump:geovars
'@prop-title Geographic variables list
'@return A BetterArray with geography-driven variables
Public Property Get GeoVars() As BetterArray: End Property

'@jump:timevars
'@prop-title Time variables list
'@return A BetterArray with time-related variables
Public Property Get TimeVars() As BetterArray: End Property

'@jump:specialvars
'@prop-title Filtered variable list by criteria
'@param first_condition String. Primary control/tag value to match.
'@param second_condition Optional String. Secondary control/tag value.
'@param conditionName Optional String. Column header to evaluate. Defaults to "Control".
'@return A BetterArray of matching variables
Public Property Get SpecialVars(ByVal first_condition As String, _
                                Optional ByVal second_condition As String = vbNullString, _
                                Optional ByVal conditionName As String = "Control") As BetterArray
End Property


'@section Preparation
'===============================================================================

'@jump:prepare
'@sub-title Prepare dictionary for designer workflows
'@param PreservedSheetNames BetterArray containing table names to preserve.
'@param geoObject ILLGeo implementation to enrich geographic variables.
'@param msgTrads ITranslationObject providing translations.
'@param tablePrefix Optional String prefix for generated table names. Defaults to "table".
Public Sub Prepare(Optional ByVal PreservedSheetNames As BetterArray, _
                   Optional ByVal geoObject As Object, _
                   Optional ByVal tablePrefix As String = "table")
End Sub

'@jump:clean
'@sub-title Clean up dictionary columns
Public Sub Clean(Optional ByVal removeAddedColumns As Boolean = False): End Sub


'@section Data Exchange
'===============================================================================

'@jump:import
'@sub-title Import dictionary data
'@param fromWksh Worksheet source worksheet.
'@param fromStartRow Long. Starting row of the source data.
'@param fromStartcol Long. Starting column of the source data.
Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
End Sub

'@jump:export
'@sub-title Export dictionary data
'@param toWkb Workbook destination workbook.
'@param exportType Optional String filter for export. Defaults to "__all__".
'@param addListObject Optional Boolean. Add a table to exported sheet. Defaults to True.
'@param Hide Optional Long sheet visibility after export. Defaults to xlSheetHidden.
Public Sub Export(ByVal toWkb As Workbook, _
                  Optional ByVal exportType As String = "__all__", _
                  Optional ByVal addListObject As Boolean = True, _
                  Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'@jump:translate
'@sub-title Translate dictionary labels
'@param TransObject ITranslationObject used to translate ranges.
Public Sub Translate(ByVal TransObject As ITranslationObject)
End Sub

