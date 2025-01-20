VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of dictionary class"

'@Folder("Dictionary")
'@ModuleDescription "Interface of dictionary class"
'@Interface
'@IgnoreModule ProcedureNotUsed

Option Explicit

'Exposed methods
Public Property Get Data() As IDataSheet: End Property
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) _ 
                              As Range
End Property

'Functions
Public Function ColumnExists(ByVal colName As String, _ 
                             Optional ByVal checkValidity As Boolean) _ 
                            As Boolean
End Function

Public Function VariableExists(ByVal varName As String) As Boolean: End Function

'Add and remove rows
Public Sub AddRows(): End Sub
Public Sub RemoveRows(): End Sub
Public Sub InsertColumn(ByVal colName As String, ByVal After As String): End Sub
Public Sub RemoveColumn(ByVal colName As String): End Sub
Public Property Get SpecialVars(ByVal first_condition As String, _
                                Optional ByVal second_condition As String = vbNullString, _
                                Optional ByVal conditionName As String = "Control") _ 
                                As BetterArray
End Property
Public Property Get ChoicesVars() As BetterArray: End Property
Public Property Get GeoVars() As BetterArray: End Property
Public Property Get TimeVars() As BetterArray: End Property


