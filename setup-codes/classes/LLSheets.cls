VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist sheets"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist sheets")

Option Explicit

Implements ILLSheets

'Instanciation
Private Type TLLSheets
    dict As ILLdictionary
    sheetRng As Range
End Type

Private this As TLLSheets
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const CLASSNAME As String = "LLSheets"
Private Const SHEETHORIZONTALNBROWS As Long = 201
Private Const SHEETNAMECOLUMN As String = "Sheet Name"
Private Const SHEETTYPECOLUMN As String = "Sheet Type"
Private Const TABLENAMECOLUMN As String = "Table Name"

Private Enum SHEETSTART
    SHEETSTARTCOLHORIZONTAL = 1
    SHEETSTARTROWHORIZONTAL = 8
    SHEETSTARTROWVERTICAL = 4
    SHEETSTARTCOLVERTICAL = 5
End Enum

Public Function Create(ByVal dict As ILLdictionary) As ILLSheets

    If (Not dict.ColumnExists(SHEETNAMECOLUMN)) Or (Not dict.ColumnExists(SHEETTYPECOLUMN)) Then _
        ThrowErrorUnableInstanciateNoSheetName

    With New LLSheets
        Set .Dictionary = dict
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLSheets
    Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    Set this.sheetRng = dict.DataRange(SHEETNAMECOLUMN, includeHeaders:=True)
End Property

Public Property Get SheetRange() As Range
    Set SheetRange = this.sheetRng
End Property

'methods
Private Function Contains(ByVal sheetName As String) As Boolean
    Dim sheetRng As Range
    Dim Exists As Boolean

    Set sheetRng = SheetRange()
    If sheetName = vbNullString Or sheetName = SHEETNAMECOLUMN Then
        Exists = False
    Else
        Exists = Not (sheetRng.Find(What:=sheetName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If

    Contains = Exists
End Function

'Row Index for table and varlists
Private Function RowIndex(ByVal sheetName As String) As Long
    Dim sheetRng As Range
    Dim rowNum As Long
    rowNum = -1
    Set sheetRng = SheetRange()
    If Contains(sheetName) Then
        rowNum = sheetRng.Find(What:=sheetName, lookAt:=xlWhole, MatchCase:=True).Row
    End If
    RowIndex = rowNum
End Function

'Get the sheet information

Public Function SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = 1) As String
    Dim sheetInfoIndex As Long
    Dim retSheetInfo As String
    Dim dict As ILLdictionary
    Dim dictWksh As Worksheet

    Set dict = Dictionary
    Set dictWksh = dict.Data.Wksh
    retSheetInfo = vbNullString

    If Not Contains(sheetName) Then Exit Function

    If (info = 1) Then
        sheetInfoIndex = dict.Data.ColumnIndex(SHEETTYPECOLUMN)
    ElseIf (info = 2) Then
        sheetInfoIndex = dict.Data.ColumnIndex(TABLENAMECOLUMN)
    Else
        Exit Function
    End If

    retSheetInfo = dictWksh.Cells(RowIndex(sheetName), sheetInfoIndex)
    SheetInfo = retSheetInfo
End Function

'List of all variables of a sheet
Private Function Varlist(ByVal sheetName As String) As BetterArray

    Dim varlistData As BetterArray
    Dim dict As ILLdictionary

    Set dict = Dictionary
    Set varlistData = New BetterArray
    If Contains(sheetName) Then Set varlistData = dict.Data.FilterData(SHEETNAMECOLUMN, sheetName, "Variable Name")

    'return the variable list
    Set Varlist = varlistData.Clone()
End Function

'Number of variables in a a sheet
Private Function NumberVars(ByVal sheetName As String) As Long
    Dim varData As BetterArray
    Set varData = New BetterArray
    Set varData = Varlist(sheetName)
    NumberVars = varData.Length
End Function

'Implement the databounds
Private Function DataBounds(ByVal sheetName As String, Optional ByVal Bound As Byte = 1) As Long

    Dim nbVars As Long

    Select Case Bound

    Case 1
        DataBounds = Switch( _
                     SheetInfo(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTROWHORIZONTAL, _
                     SheetInfo(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTROWVERTICAL _
                                                              )
    Case 2
        nbVars = NumberVars(sheetName)
        DataBounds = Switch( _
                     SheetInfo(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTROWHORIZONTAL + SHEETHORIZONTALNBROWS, _
                     SheetInfo(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTROWVERTICAL + nbVars - 1)
    Case 3
        DataBounds = IIf(SheetInfo(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTCOLHORIZONTAL, SHEETSTARTCOLVERTICAL)
    Case 4
        nbVars = NumberVars(sheetName)
        DataBounds = Switch( _
                     SheetInfo(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTCOLHORIZONTAL + nbVars - 1, _
                     SheetInfo(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTCOLVERTICAL)
    End Select

End Function

'Test if a sheet contains a type of control (list_auto by default)
Private Function ContainsControl(ByVal sheetName As String, _
                                 Optional ByVal control As String = "list_auto") As Boolean
    'Check if a sheet contains a list auto
    Dim controlVar As BetterArray
    Dim varData As BetterArray
    Dim dict As ILLdictionary
    Dim counter As Long

    Set controlVar = New BetterArray
    Set varData = New BetterArray
    Set dict = Dictionary()

    Dim exist As Boolean                         'Test if there is a listauto var in the sheet (at least one)

    exist = False

    Set controlVar = dict.Data.FilterData("Control", control, "Variable Name")
    Set varData = Varlist(sheetName)

    If controlVar.Length <> 0 Then
        For counter = controlVar.LowerBound To controlVar.UpperBound
            If varData.Includes(controlVar.Item(counter)) Then
                exist = True
                Exit For
            End If
        Next
    End If

    ContainsControl = exist
End Function

'Variable Address
Private Function VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String
    'Return the variable address
    Dim varAdd As String
    Dim varIndex As Long
    Dim rstart As Long
    Dim cstart As Long
    Dim dict As ILLdictionary
    Dim sheetName As String
    Dim dictWksh As Worksheet
    Dim vars As ILLVariables
    Dim prefix As String                         'Prefix to variable address

    Set dict = Dictionary()
    Set dictWksh = dict.Data.Wksh()
    Set vars = LLVariables.Create(dict)

    varIndex = vars.Index(varName)
    sheetName = vars.Value(colName:=SHEETNAMECOLUMN, varName:=varName)

    prefix = vbNullString
    
    If (onSheet <> sheetName) Then
        prefix = "'" & sheetName & "'!"
    End If
  
    rstart = DataBounds(sheetName, 1)
    cstart = DataBounds(sheetName, 3)

    If Contains(sheetName) Then
        Select Case SheetInfo(sheetName)
        Case SHEETTYPEHORIZONTAL
            varAdd = prefix & dictWksh.Cells(rstart + 1, varIndex).Address(Rowabsolute:=False)
        Case SHEETTYPEVERTICAL
            varAdd = prefix & dictWksh.Cells(varIndex, cstart).Address(Rowabsolute:=True, ColumnAbsolute:=True)
        End Select
    End If
  
    VariableAddress = varAdd
End Function

'Errors
Private Sub ThrowErrorUnableInstanciateNoSheetName()
    Err.Raise VBA.vbObjectError + 513, CLASSNAME, "Unable to create an instance of the LLSheets variable, sheet name and sheet types columns do not exists"
End Sub

Private Sub ThrowErrorUnableInstanciateNotPrepared()
    Err.Raise VBA.vbObjectError + 514, CLASSNAME, "Unable to create an instance of the LLSheets variable: The dictionary is not prepared"
End Sub

'Interface ---------------------------------------------------------------------
Private Property Get ILLSheets_Dictionary() As ILLdictionary
    Set ILLSheets_Dictionary = Dictionary()
End Property

Private Function ILLSheets_SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = 1) As String
    ILLSheets_SheetInfo = SheetInfo(sheetName, info)
End Function

Private Function ILLSheets_Contains(ByVal sheetName As String) As Boolean
    ILLSheets_Contains = Contains(sheetName)
End Function

Private Function ILLSheets_DataBounds(ByVal sheetName As String, _
                                     Optional ByVal Bound As Byte = 1) As Long
    ILLSheets_DataBounds = DataBounds(sheetName, Bound)
End Function

Private Function ILLSheets_ContainsControl(ByVal sheetName As String, Optional ByVal control As String = "list_auto") As Boolean
    ILLSheets_ContainsControl = ContainsControl(sheetName, control)
End Function

'Variable Address
Private Function ILLSheets_VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String
    ILLSheets_VariableAddress = VariableAddress(varName, onSheet)
End Function

'Row index of the sheet in the dictionary
Private Function ILLSheets_RowIndex(ByVal sheetName As String) As Long
    ILLSheets_RowIndex = RowIndex(sheetName)
End Function

'Number of variables in the SheetS
Private Function ILLSheets_NumberOfVars(ByVal sheetName As String) As Long
    ILLSheets_NumberOfVars = NumberVars(sheetName)
End Function


