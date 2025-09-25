VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Variable class")
'@depends Checking
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Properties
Public Property Get Dictionary() As ILLdictionary: End Property


'Exposed methods
Public Function Value(ByVal colName As String, ByVal varName As String) As String
End Function

'column or row index in the sheet of one variable
Public Function Index(ByVal varName As String) As Long
End Function

'Case-sensitive check for existence of a variable name. Set matchCase:=False to ignore letter casing.
Public Function Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
End Function

'Change the value of one column for one variable in the dictionary (use carefully: it modifies setup data)
Public Sub SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
End Sub

Public Property Get CellRange(ByVal colName As String, ByVal varName As String) As Range
End Property

'Return the list of variable names as a BetterArray snapshot.
Public Function VariableNames() As BetterArray
End Function

'Retrieve the sheet storing the variable (case-sensitive by default).
Public Function SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
End Function

'Retrieve the dictionary table assigned to a variable.
Public Function TableName(ByVal varName As String) As String
End Function

'Retrieve the control type stored in the dictionary for the variable.
Public Function ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
End Function

'Invalidate cached ranges and column lookups when the dictionary structure changes.
Public Sub InvalidateCaches()
End Sub

'Checking information ---------------------------------------------------------
Public Function HasCheckings() As Boolean
End Function

Public Property Get CheckingValues() As IChecking
End Property
