VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Linelist sheets")
'@IgnoreModule

Option Explicit

Implements ILLSheets

'Instanciation
Private Type TLLSheets
  dict As ILLdictionary
  sheetRng As Range
End Type

Private this as TLLSheets
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const CLASSNAME As String = "LLSheets"
Private Const SHEETSNAME As String = "sheet name"
Private Const SHEETSTYPE As String = "sheet type"
Private Const SHEETHORIZONTALNBROWS As Long = 200

Private Enum SHEETSTART
  SHEETSTARTCOLHORIZONTAL = 1
  SHEETSTARTROWHORIZONTAL = 9 'Where headers of the Listobject are in the table
  SHEETSTARTROWVERTICAL = 4
  SHEETSTARTCOLVERTICAL = 5 'Where the data for the list object is stored
End Enum

Public Function Create(ByVal dict As ILLdictionary) As ILLSheets

  If (Not dict.ColumnExists(SHEETSNAME)) Or (Not dict.ColumnExists(SHEETSTYPE)) Then ThrowErrorUnableInstanciateNoSheetName
  If Not dict.Prepared Then ThrowErrorUnableInstanciateNotPrepared

  With New LLSheets
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

Private Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    Set this.sheetRng = dict.DataRange(SHEETSNAME)
End Property

Private Property SheetRange() As Range
    Set SheetRange = this.sheetRng
End Property

'methods
Private Function Contains(ByVal sheetName As String) As Boolean
  Dim sheetRng As Range
  Dim Exists As Boolean

  Set sheetRng = SheetRange()
  If sheetName = vbNullString Then
    Exists = False
  Else
    Exist = Not (sheetRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
  End If

  Contains = Exist
End Function

'Row Index for table and varlists
Private Function RowIndex(ByVal sheetName As String) As Long
  Dim sheetRng As Range
  Dim rowNum As Long
  rowNum = -1
  Set sheetRng = SheetRange()
  If Contains(sheetName) Then
      rowNum = sheetRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True).Row
  End If
  RowIndex = rowNum
End Function

''Sheet type of one sheet Name
Private Function SheetType(ByVal sheetName As String) As String

  Dim sheetTypeIndex As Long
  Dim retSheetType As String
  Dim dict As ILLdictionary
  Dim dictWksh As Worksheet

  Set dict = Dictionary
  Set dictWksh = dict.Wksh
  retSheetType = vbNullString

  sheetTypeIndex = dict.Data.ColumnIndex(SHEETSTYPE)

  If Contains(sheetName) Then retSheetType = dictWksh.Cells(RowIndex(sheetName), sheetTypeIndex)

  SheetType = retSheetType
End Function

'List of all variables of a sheet
Private Function Varlist(ByVal sheetName As String) As BetterArray

  Dim varlistData As BetterArray
  Dim dict As ILLdictionary

  Set dict = Dictionary
  Set varlistData = New BetterArray
  If Contains(sheetName) Then  Set varlistData = dict.Data.Filter(SHEETSNAME, sheetName, "variable name")

  'return the variable list
  Set VarList = varlistData.Clone()
End Function


'Table of one variable
Private Function Table(ByVal sheetName As String) As String
  Dim tableName As String

  Dim dict As ILLdictionary
  Dim dictWksh As Worksheet

  Set dict = Dictionary
  Set dictWksh = dict.Wksh
  tableName = vbNullString


  If Contains(sheetName) And dict.ColumnExists("table name") Then
    tableName = dictWksh.Cells(RowIndex(sheetName), dict.Data.ColumnIndex("table name"))
  End If

  Table = tableName
End Function

'Number of variables in a a sheet
Private Function NumberVars(ByVal sheetName As String) As Long
    Dim varData As BetterArray
    Set varData = New BetterArray
    Set varData = VarList(sheetName)
    NumberVars = varData.Length
End Function

'Implement the databounds
Private Function DataBounds(ByVal sheetName As String, Optional ByVal SheetBound As SheetBound)

    Dim bound As Long
    Dim nbVars As Long

    Select Case SheetBound

      Case RowStart
        bound = Switch( _
        SheetType(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTROWHORIZONTAL, _
        SheetType(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTROWVERTICAL _
        )
      Case RowEnd
        nbVars = NumberVars(sheetName)
        bound = Switch( _
          SheetType(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTROWHORIZONTAL + SHEETHORIZONTALNBROWS, _
          SheetType(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTROWVERTICAL + nbVars - 1 )
      Case ColStart
        bound = IIf(SheetType(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTCOLHORIZONTAL, SHEETSTARTCOLVERTICAL)
      Case ColEnd
        nbVars = NumberVars(sheetName)
        bound = Switch(
          SheetType(sheetName) = SHEETTYPEHORIZONTAL, SHEETSTARTCOLHORIZONTAL + nbVars - 1, _ 
          SheetType(sheetName) = SHEETTYPEVERTICAL, SHEETSTARTCOLVERTICAL)
    End Select

    DataBounds = bound
End Function


'ContainsListAuto
Private Function ContainsListAuto(ByVal sheetName As String) As Boolean
    'Check if a sheet contains a list auto
    Dim listAutoVar As BetterArray
    Dim varData As BetterArray
    Dim dict As ILLdictionary
    Dim counter As Long

    Set listAutoVar = New BetterArray
    Set varData = New BetterArray
    Set dict = Dictionary()

    Dim Exist As Boolean 'Test if there is a listauto var in the sheet (at least one)

    Exist = False

    Set listAutoVar = dict.Data.Filter("control", "list_auto", "variable name")
    Set varData = VarList(sheetName)

    If listAutoVar.Length <> 0 Then
      For counter = listAutoVar.LowerBound To listAutoVar.UpperBound
        If Not Exist And varData.Includes(listAutoVar.Item(counter)) Then
            Exist = True
            Exit For
        End If
      Next
    End If

    ContainsListAuto = Exist
End Function

'Variable Address
Private Function VariableAddress(Byval sheetName As String, ByVal varName As String, _ 
                                 Optional ByVal useTable As Boolean = False) As String

  'Return the variable address
  Dim varAdd As String
  Dim varIndex As Long
  Dim rstart As Long
  Dim cstart As Long
  Dim dict As ILLdictionary
  Dim dictWksh As Worksheet
  Dim vars As ILLVariables

  Set dict = Dictionary()
  Set dictWksh = dict.Wksh()
  set vars = LLVariables.Create(dict)

  varIndex = vars.Index(varName)
  rstart = DataBounds(sheetName, RowStart)
  cstart = DataBounds(sheetName, ColStart)

  Select Case SheetType(sheetName)

    Case SHEETTYPEHORIZONTAL

    Case SHEETTYPEVERTICAL

  End Select

End Function


'Errors
Private Sub ThrowErrorUnableInstanciateNoSheetName()
    Err.Raise VBA.vbObjectError + 513, CLASSNAME, "Unable to create an instance of the class variable, sheet name and sheet types columns do not exists"
End Sub

Private Sub ThrowErrorUnableInstanciateNotPrepared()
    Err.Raise VBA.vbObjectError + 514, CLASSNAME, "Unable to create an instance of the class variable: The dictionary is not prepared"
End Sub


'Interface ---------------------------------------------------------------------
Public Function ILLSheets_SheetType(ByVal sheetName As String) As String
  ILLSheets_SheetType =  SheetType(sheetName)
End Function

Public Function ILLSheets_Contains(ByVal sheetName As String) As Boolean
  ILLSheets_Contains =  Contains(sheetName)
End Function

Public Function ILLSheets_Table(ByVal sheetName As String) As String
  ILLSheets_Table =  Table(sheetName)
End Function

Public Function ILLSheets_DataBounds(ByVal sheetName As String, _
                                     Optional ByVal Bound As SheetBound) As Long
  ILLSheets_DataBounds =  DataBounds(sheetName, start, row)
End Function

Public Function ILLSheets_VarList(ByVal sheetName As String) As BetterArray
  Set ILLSheets_VarList =  VarList(sheetName).Clone()
End Function


Public Function ILLSheets_ContainsListAuto(ByVal sheetName As String) As Boolean
  ILLSheets_ContainsListAuto =  ContainsListAuto(sheetName)
End Function
