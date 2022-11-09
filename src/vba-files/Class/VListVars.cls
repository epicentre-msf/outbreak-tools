VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VListVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Vlist sheet")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements IVListVars

Private Type TVListVars
    data As ILinelistSpecs
    drop As IDropdownLists
    sh As Worksheet
    varName As String
    inputRange As Range
End Type

Private this As TVListVars
Private Const CLASSNAME As String = "VListVars"
Private Const STARTCOL As Long = 5            'Where the variables should start in the output worksheet

Public Function Create(ByVal varName As String, ByVal outputsh As Worksheet, _
                       ByVal lData As ILinelistSpecs, ByVal drop As IDropdownLists) As IVListVars
    Dim dict As ILLdictionary

    If outputsh Is Nothing Then ThrowError "Worksheet does not exists"
    If lData Is Nothing Then ThrowError "No linelist specifications found"
    Set dict = lData.Dictionary()
    If Not dict.VariableExists(varName) Then ThrowError "Variable" & varName & "Does not exists!"

    With New VListVars
        Set .linelistData = lData
        Set .Wksh = outputsh
        Set .Dropdown = drop
        .ActualVariable = varName
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IHListVars
    Set Self = Me
End Property

Public Property Get linelistData() As ILinelistSpecs
    Set linelistData = this.data
End Property

Public Property Set linelistData(ByVal data As ILinelistSpecs)
    Set this.data = data
End Property

Public Property Get ActualVariable() As String
    ActualVariable = this.varName
End Property

Public Property Let ActualVariable(ByVal varName As String)
    this.varName = varName
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Get Dropdown() As IDropdownLists
    Set Dropdown = this.drop
End Property

Public Property Set Dropdown(ByVal drop As IDropdownLists)
    Set this.drop = drop
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Private Property Get Dictionary() As ILLdictionary
  Dim lData As ILinelistSpecs
  Set lData = linelistData()
  Set Dictionary = lData.Dictionary
End Property

Private Property Get vars() As ILLVariables
    Set vars = LLVariables.Create(Dictionary())
End Property

Private Property Get ValueOf(ByVal colName As String) As String
  Dim varObject As ILLVariables
  Set varObject = vars()
  ValueOf = varObject.Value(colName:=colName, varName:=ActualVariable)
End Property

Private Property Get VarRange() As Range         'Where to input informations on the variable
  Dim startRow As Long
  Dim sh As Worksheet
  Dim rng As Range
  If this.inputRange Is Nothing Then
    Set sh = Wksh()
    With sh
      startRow = ValueOf("column index")
      Set VarRange = .Cells(startRow, STARTCOL)
      Set this.inputRange = .Cells(startRow, STARTCOL)
    End With
  Else
    Set VarRange = this.inputRange
  End If
End Property


'Type
Private Sub AddType()
    Dim varType As String
    Dim cellRng As Range
    Dim nbDecimals As Integer
    Dim zerosDecimals As String

    Set cellRng = VarRange()

    varType = LCase(Application.WorksheetFunction.Trim(ValueOf("type")))

    'Number of decimals
    If InStr(1, varType, "decimal") > 0 Then
        nbDecimals = CInt(Replace(varType, "decimal", vbNullString))
        zerosDecimals = String(nbDecimals, "0")
    End If

    Select Case varType
    Case "text"
        cellRng.NumberFormat = "@"
    Case "integer"
        cellRng.NumberFormat = "0"
    Case "date"
        cellRng.NumberFormat = "d-mmm-yyy"
    Case "decimal"
        cellRng.NumberFormat = "0." & zerosDecimals
    Case Else
        cellRng.NumberFormat = "@"
    End Select
End Sub

'Failed formula
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim sh As Worksheet

    Set sh = Wksh()
    With sh
        .Cells(1, 1).Clear
        On Error GoTo Fail
        Application.DisplayAlerts = False

        'Test the formula on one cell
        .Cells(1, 1).formula = "= " & frmName

        .Cells(1, 1).Clear
        FailedFormula = False
        Exit Function
    End With
    
    Application.DisplayAlerts = True
Fail:
    FailedFormula = True
    Application.DisplayAlerts = True
End Function

'Formula
Private Sub AddFormula()
  Dim varForm As String
  Dim excelForm As String
  Dim cellRng As Range
  Dim lData As ILinelistSpecs
  Dim varFormData As IFormulaData
  Dim dict As ILLdictionary
  Dim varFormObject As IFormulas
  Dim DesignFormat As ILLFormat
  Dim varCaseWhen As ICaseWhen
  Dim varControl As String
  varControl = ValueOf("control")
  If (varControl <> "formula") And (varControl <> "case_when") Then Exit Sub
  
  Set lData = linelistData()
  Set dict = lData.Dictionary()
  Set varFormData = lData.FormDataObject()
  Set DesignFormat = lData.DesignFormat()
  varForm = ValueOf("control details")
  Set cellRng = VarRange()
  'For case_when parse the case_when before
  If varControl = "case_when" Then
      Set varCaseWhen = CaseWhen.Create(varForm)
      varForm = varCaseWhen.parsedFormula()
  End If
  Set varFormObject = Formulas.Create(dict, varFormData, varForm)
  If varFormObject.valid(formulaType:="linelist") Then
      excelForm = varFormObject.ParsedLinelistFormula()
      'Test if the formula is correct and parse it
      If FailedFormula(excelForm) Then
         excelForm = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
      End If
      cellRng.formula = "= " & excelForm
      cellRng.Locked = True
  End If
  'Apply format of the calculated range
  DesignFormat.ApplyFormat cellRng, scope:=HListCalculatedFormulaCell
End Sub

'Main Label
Private Sub AddLabel(Optional ByVal notevisibility As Boolean = False)
  Dim varMainLab As String
  Dim varName As String
  Dim varNote As String
  Dim varSublab As String
  Dim varStatus As String
  Dim varListAutoOrigin As String
  Dim varControl As String
  Dim cellRng As Range
  Dim trans As ITranslation
  Dim lData As ILinelistSpecs
  Dim calculated As String
  Dim mandatory As String
  Dim mainLabLength As Long
  Dim design As ILLFormat


  Set cellRng = VarRange()
  Set lData = linelistData
  Set trans = lData.TransObject()
  Set design = lData.DesignFormat()

  varMainLab = ValueOf("main label")
  varName = ValueOf("variable name")
  varNote = ValueOf("note")
  varSublab = ValueOf("sub label")
  varControl = ValueOf("control")
  varStatus = ValueOf("status")
  calculated = trans.TranslatedValue("MSG_Calculated")
  mandatory = trans.TranslatedValue("MSG_Mandatory")

  'sublabel
  If (varControl = "formula") Or (varControl = "case_when") Then
      varSublab = varSublab & Chr(10) & calculated
  End If

  mainLabLength = Len(varMainLab)
  'Add subLabel to main label
  If (varSublab <> vbNullString) Then
     varMainLab = varMainLab & Chr(10) & varSublab
  End If
  'Status to add to Notes
  If varStatus = "mandatory" Then varNote = mandatory & Chr(10) & varNote
  'variable main label and sub lable
  cellRng.Offset(, -1).Value = varMainLab
  
  design.ApplyFormat cellRng.Offset(, -1).Characters(start:=1, Length:=mainLabLength), _
      scope:=VListMainLab   'format of the main label
  
  If (varSublab <> vbNullString) Then
    design.ApplyFormat cellRng.Offset(, -1).Characters(start:=mainLabLength + 1), _
      scope:=VListSublab   'format of the sub label
  End If
  
  'variable name
  cellRng.Name = varName
  
  'variable note
  If varNote <> vbNullString Then
      cellRng.AddComment
      cellRng.Comment.text text:=varNote
      cellRng.Comment.Visible = notevisibility
  End If

  'Hidden Row
  If varStatus = "hidden" Then cellRng.EntireRow.Hidden = True
End Sub


'Add the choices
Private Sub AddChoices()
    Dim varControl As String
    Dim varName As String
    Dim varAlert As String
    Dim varMessage As String

    Dim lData As ILinelistSpecs
    Dim cat As BetterArray
    Dim drop As IDropdownLists
    Dim cellRng As Range

    Set lData = linelistData()
    Set drop = Dropdown()
    Set cat = New BetterArray

    varName = ActualVariable()
    varControl = ValueOf("control")

    If (varControl <> "choices") Then Exit Sub

    Set cat = lData.Categories(varName)
    Set cellRng = VarRange()
    
    'Add dropdown
    drop.Add cat, varName
    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")

    'Set dropdown validations
    drop.SetValidation cellRng:=cellRng, ListName:=varName, _
                       alertType:=varAlert, message:=varMessage
End Sub


'Write one variable information in the corresponding cell.

Private Sub WriteInfo()

    'Add the label
    AddLabel

    'Formulas
    AddFormula

    'Types
    AddType

    'Choices
    AddChoices
End Sub


'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Public Sub IHListVars_WriteInfo()
    WriteInfo
End Sub

Public Property Get IHListVars_ValueOf(ByVal colName As String) As String
    IHListVars_ValueOf = ValueOf(colName)
End Property
