VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist variables class"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist variables class")

'This is a modified version of LLvariables without need for the Dictionary
'to be prepared

Option Explicit

Implements ILLVariables

Private Type TVariable
    dict As ILLdictionary
    varRng As Range
End Type

Private this As TVariable
Private Const VARIABLENAME As String = "Variable Name"
Private Const CLASSNAME As String = "LLVariables"

Public Function Create(ByVal dict As ILLdictionary) As ILLVariables

    If Not dict.ColumnExists(VARIABLENAME) Then _ 
     ThrowErrorUnableInstanciateNoVarName

    With New LLVariables
        Set .Dictionary = dict
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLVariables
    Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    Set this.varRng = dict.DataRange(VARIABLENAME)
End Property

Friend Property Get VarRange() As Range
    Set VarRange = this.varRng
End Property

'Functions ---------------------------------------------------------------------

'Test if a variable is present
Private Function Contains(ByVal varName As String) As Boolean
    Dim varRng As Range
    Dim Exist As Boolean
    Dim searchValue As String

    'Removing all wilcards in search
    searchValue = Replace(varName, "*", "~*")
    searchValue = Replace(searchValue, "?", "~?")
    searchValue = Replace(searchValue, "!", "~!")
    searchValue = Replace(searchValue, "[", "~[")
    searchValue = Replace(searchValue, "]", "~]")
    searchValue = Replace(searchValue, "#", "~#")

    Set varRng = VarRange()
    If (searchValue = vbNullString) Then
        Exist = False
    Else
        On Error Resume Next
        Exist = Not (varRng.Find(What:=searchValue, _ 
                     LookAt:=xlWhole, MatchCase:=True) Is Nothing)
        On Error GoTo 0
    End If

    Contains = Exist
End Function

'Find the row index of one variable
Private Function RowIndex(ByVal varName As String) As Long
    Dim varRng As Range
    Dim rowNum As Long
    rowNum = -1
    Set varRng = VarRange()
    If Contains(varName) Then
        rowNum = varRng.Find(What:=varName, lookAt:=xlWhole, MatchCase:=True).Row
    End If
    RowIndex = rowNum
End Function

'Value of one column for one variable
Private Function Value(ByVal colName As String, ByVal varName As String) As String

    'Row and columns to return
    Dim rowNum As Long
    Dim colNum As Long
    Dim retValue As String

    Dim dict As ILLdictionary
    Dim Wksh As Worksheet

    Set dict = Dictionary()
    Set Wksh = dict.Data.Wksh

    'return value
    retValue = vbNullString

    If Contains(varName) And dict.ColumnExists(colName) Then

        rowNum = RowIndex(varName)
        colNum = dict.Data.ColumnIndex(colName)
        retValue = Wksh.Cells(rowNum, colNum).Value

    End If

    Value = retValue
End Function

Private Sub SetValue(ByVal varName As String, ByVal colName As String, _ 
                     ByVal newValue As String, _ 
                    Optional ByVal onEmpty As Boolean = False)

    'Row and columns to return
    Dim rowNum As Long
    Dim colNum As Long
    Dim Update As Boolean

    Dim dict As ILLdictionary
    Dim Wksh As Worksheet

    Set dict = Dictionary()
    Set Wksh = dict.Data.Wksh

    If Contains(varName) And dict.ColumnExists(colName) Then
        Update = True
        
        rowNum = RowIndex(varName)
        colNum = dict.Data.ColumnIndex(colName)
        If onEmpty Then Update = IsEmpty(Wksh.Cells(rowNum, colNum))
        If Update Then Wksh.Cells(rowNum, colNum).Value = newValue
    End If

End Sub


'Errors ------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciateNoVarName()
    Err.Raise VBA.vbObjectError + 513, CLASSNAME, _ 
     "Unable to create an instance of the class variable, variable name column does not exists"
End Sub


'Interface ---------------------------------------------------------------------

'Properties
Private Property Get ILLVariables_Dictionary() As ILLdictionary
    Set ILLVariables_Dictionary = Dictionary()
End Property

'Exposed methods
Private Function ILLVariables_Value(ByVal colName As String, _ 
                                    ByVal varName As String) As String
    ILLVariables_Value = Value(colName, varName)
End Function


Private Function ILLVariables_Contains(ByVal varName As String) As Boolean
    ILLVariables_Contains = Contains(varName)
End Function

Private Sub ILLVariables_SetValue(ByVal varName As String, _ 
                                  ByVal colName As String, _ 
                                  ByVal newValue As String, _ 
                                  Optional ByVal onEmpty As Boolean = False)
    SetValue varName, colName, newValue, onEmpty
End Sub

