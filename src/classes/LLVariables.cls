VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Linelist Variable Class")
'@IgnoreModule

Option Explicit

Implements ILLVariables

Private Type TVariable
    dict As ILLdictionary
    varRng As Range
End Type

Private this As TVariable
Private Const VARIABLENAME As String = "variable name"
Private Const INDEXHEADER  As String = "column index"
Private Const CLASSNAME As String = "LLVariables"

Public Function Create(ByVal dict As ILLdictionary) As ILLVariables

    If Not dict.ColumnExists(VARIABLENAME) Then ThrowErrorUnableInstanciateNoVarName
    If Not dict.Prepared Then ThrowErrorUnableInstanciateNotPrepared

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

'Functions --------------------------------------------------------------------------

'Test if a variable is present
Private Function Contains(ByVal varName As String) As Boolean
    Dim varRng As Range
    Dim Exist As Boolean
  
    Set varRng = VarRange()
    If varName = vbNullString Then
        Exist = False
    Else
        Exist = Not (varRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
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
        rowNum = varRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True).Row
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

Private Sub SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)

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

'Index of a column or a Row of a variable in a worksheet
Private Function index(ByVal varName As String) As Long
    Dim sheetIndex As String
    sheetIndex = Value(colName:=INDEXHEADER, varName:=varName)
    index = IIf(sheetIndex = vbNullString, -1, CLng(sheetIndex))
End Function

'Errors ------------------------------------------------------------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciateNoVarName()
    Err.Raise VBA.vbObjectError + 513, CLASSNAME, "Unable to create an instance of the class variable, variable name column does not exists"
End Sub

Private Sub ThrowErrorUnableInstanciateNotPrepared()
    Err.Raise VBA.vbObjectError + 514, CLASSNAME, "Unable to create an instance of the class variable: The dictionary is not prepared"
End Sub

'Interface ---------------------------------------------------------------------------------------------------------------------------------

'Properties
Public Property Get ILLVariables_Dictionary() As ILLdictionary
    Set ILLVariables_Dictionary = Dictionary
End Property

'Exposed methods
Public Function ILLVariables_Value(ByVal colName As String, ByVal varName As String) As String
    ILLVariables_Value = Value(colName, varName)
End Function

'column or row index in the sheet of one variable
Public Function ILLVariables_Index(ByVal varName As String) As Long
    ILLVariables_Index = index(varName)
End Function

Public Function ILLVariables_Contains(ByVal varName As String) As Boolean
    ILLVariables_Contains = Contains(varName)
End Function

Public Sub ILLVariables_SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
    SetValue varName, colName, newValue, onEmpty
End Sub

