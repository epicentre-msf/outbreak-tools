VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "CRFVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "CRF Variables"

'@Folder("Linelist")
'@ModuleDescription("CRF Variables")
'@IgnoreModule  UseMeaningfulName, HungarianNotation

Option Explicit
Implements ICRFVars

Private Type TCRFVar
  ll As ILinelist
  varName As String
  inputRange As Range
End Type

Private this As TCRFVar
Private const CLASSNAME As String = "CRFVariables"


Public Function Create(ByVal varName As String, ll As ILinelist) As ICRFVars
  
  CheckRequirements varName, ll

  With New CRFVars
    .ActualVariable = varName
    Set .LinelistObject = ll
    Set Create = .Self
  End With

End Function

Public Property Get Self() As ICRFVars
  Set Self = Me
End Property

Public Property Let ActualVariable(ByVal varName As String)
  this.varName = varName
End Property

Public Property Get ActualVariable() As String
  ActualVariable = this.varName
End Property

Public Property Get LinelistObject() As ILinelist
  Set LinelistObject = this.ll
End Property

Public Property Set LinelistObject(ByVal ll As ILinelist)
  Set this.ll = ll
End Property

'Get the linelist specifications, ditionary, the variables, and the design
'transitory objects

Private Property Get LinelistData() As ILinelistSpecs
  Dim ll As ILinelist
  Set ll = LinelistObject
  Set LinelistData = ll.LinelistData()
End Property


Private Property Get Dictionary() As ILLdictionary
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set Dictionary = lData.Dictionary()
End Property

Private Property Get DesignFormat() As ILLFormat
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set DesignFormat = lData.DesignFormat()
End Property

'variables object
Private Property Get Vars() As ILLVariables
    Dim dict As ILLdictionary
    Set dict = Dictionary()
    Set Vars = LLVariables.Create(dict)
End Property

'Get value for one column in the dictionary
Private Property Get ValueOf(ByVal colName As String) As String
  Dim varobj As ILLVariables
  Dim varName As String

  Set varobj = Vars()
  varName = ActualVariable()

  ValueOf = varobj.Value(colName:=colName, varName:=varName)
End Property

'Busy app
Private Sub BusyApp()
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
  Application.EnableAnimations = False
End Sub

'Initialize the CRF worksheet if required
Private Sub InitializeCRFSheet()
  Const SHCRFPREFIX As String = "crf_"

  Dim ll As ILinelist
  Dim sh As Worksheet
  Dim design As ILlFormat
  Dim sheetName As String
  Dim tabName As String

  Set ll = LinelistObject()
  Set design = DesignFormat()
  sheetName = ValueOf("sheet name")
  'add the prefix for the crf on the name of the table
  tabName = SHCRFPREFIX & ValueOf("table name")
  
  'Create the crf worksheet
  ll.AddOutputSheet sheetName, xlSheetHidden, sheetScope:=3

  'Format the crf worksheet
  Set sh = ll.CRFWksh(sheetName)

  design.ApplyFormat sh, scope:=HListCRFSheet

  'Page laoyout on the crf worksheet
  On Error Resume Next
    sh.Activate
    BusyApp
    ActiveWindow.View = xlPageLayoutView
  On Error GoTo 0

  'add anchor for the crf worksheet
  sh.Cells(1, 3).Value = "HList CRF"
  sh.Cells(1, 4).Value = tabName
  design.ApplyFormat sh.Cells(1, 3), scope:=LinelistHiddenCell
  design.ApplyFormat sh.Cells(1, 4), scope:=LinelistHiddenCell
End Sub

Private Property Get Wksh() As Worksheet
  Dim ll As ILinelist
  Dim sheetName As String
  Dim sh As Worksheet

  Set ll = LinelistObject()
  sheetName = ValueOf("sheet name")

  'Test if the worksheet exists, and if not, create it
  On Error Resume Next
    Set sh = ll.CRFWksh(sheetName)
  On Error GoTo 0

  'If the worksheet does not exists, then initialize it.
  If (sh Is Nothing) Then 
    InitializeCRFSheet
    Set sh = ll.CRFWksh(sheetName)
  End If

  Set Wksh = sh
End Property

Private Property Get CRFRange() As Range

  Dim sh As Worksheet
  Dim crfLine As Long

  If (this.inputRange Is Nothing) Then
    
    Set sh = Wksh()
    
    On Error Resume Next
    crfLine = CLng(ValueOf("crf index"))
    On Error GoTo 0

    If (crfLine = 0) Then 
      ThrowError "Error when creating CRF Range: Unable to find the crf index for " & _ 
                 ActualVariable
    End If
    
    Set this.inputRange = sh.Cells(crfLine, 1)
  End If

  Set CRFRange = this.inputRange
End Property

'Add Label on a variable

Private Sub AddLabel()

  Dim sh As Worksheet
  Dim crfRng As Range
  Dim design As ILLFormat
  Dim crfLine As Long
  Dim varMainLab As String
  Dim varSubLab As String
  Dim varControl As String
  Dim varStatus As String
  Dim mainLabLength As Long
  Dim hiddenRow As Boolean
  Dim tabName As String

  Set sh = Wksh()
  Set design = DesignFormat()

  varMainLab = ValueOf("main label")
  varSubLab = ValueOf("sub label")
  varControl = ValueOf("control")
  varStatus = ValueOf("status")
  tabName = ValueOf("table name")

  'Make sure the line exists
  On Error Resume Next
  crfLine = CLng(ValueOf("crf index"))
  On Error GoTo 0
  'If crfLine = 0 Then Exit Sub
  
  'Get the crf range
  Set crfRng = sh.Cells(crfLine, 1)

  'On the first cell of a crf worksheet, add an anchor
  If crfLine = 6 Then crfRng.Name = tabName & "_CRFSTART"

  mainLabLength = Len(varMainLab)

  If (varSubLab <> vbNullString) Then
    varMainLab = varMainLab & chr(10) & varSubLab
  End If

  'writing the main label and formatting
  crfRng.Value = varMainLab
  design.ApplyFormat crfRng.Characters(Start:=1, Length:=mainLabLength), scope:=HListMainLab                      
  design.ApplyFormat crfRng.Characters(Start:=mainLabLength + 1), scope:=HListSublab

  'Hidding the columns not to put in crfs
  hiddenRow = (varStatus = "hidden" Or varStatus = "optional, hidden") Or _ 
              (varControl = "formula" Or varControl = "choice_formula" Or _ 
               varControl = "case_when")

  crfRng.EntireRow.HIDDEN = hiddenRow
End Sub


Private Sub AddType()

  Dim crfRng As Range
  Dim varTypes As String
  Dim varFormat As String
  Dim varControl As String
  Dim scope As FormatScopes
  Dim nbDecimals As Integer
  Dim design As ILLFormat


  Set design = DesignFormat()
  Set crfRng = CRFRange()

  varTypes = LCase(ValueOf("variable type"))
  varFormat = LCase(ValueOf("variable format"))
  varFormat = Application.WorksheetFunction.Trim(varFormat)
  varControl = ValueOf("control")
  nbDecimals = 0

  If (varControl = "choice_manual") Then Exit Sub

  Select Case varTypes

  Case "text"

    'Add Lines only for text variables (not choices)
    scope = CRFText

  Case "date"

    'formating the date variable
    scope = IIf(varFormat = vbNullString, CRFDate, CRFFormatCustom)

  Case "decimal", "integer"

    If (varFormat = vbNullString) Then
      
      scope = IIf(varTypes = "decimal", CRFDecimal, CRFInteger)
    
    ElseIf (varFormat = "integer") Or (varFormat = "round0") Then
      
      scope = CRFInteger
    
    ElseIf (varFormat = "percentage0") Then

      scope = CRFPercentage
    
    ElseIf (InStr(1, varFormat, "round") > 0) Then
      
      On Error Resume Next
      nbDecimals = CInt(Replace(varFormat, "round", vbNullString))
      On Error GoTo 0

      scope = CRFDecimal

    ElseIf (InStr(1, varFormat, "percentage") > 0) Then
      
      On Error Resume Next
      nbDecimals = CInt(Replace(varFormat, "percentage", vbNullString))
      On Error GoTo 0
      
      scope = CRFPercentage
    Else
      scope = CRFFormatCustom
    End If
  End Select

  'Apply required format
  design.ApplyFormat crfRng, scope:=scope, nbDecimals:=nbDecimals
End Sub


Private Sub AddChoices()

  Dim choiceHeader As String
  Dim lData As ILinelistSpecs
  Dim cat As BetterArray
  Dim crfRng As Range
  Dim cellRng As Range
  Dim counter As Long
  Dim varControl As String
  Dim design As ILLFormat
  Dim varName As String

  varControl = ValueOf("control")
  
  If (varControl <> "choice_manual") Then Exit Sub
  
  Set lData = LinelistData()
  Set crfRng = CRFRange()
  Set design = DesignFormat()
  Set cat = New BetterArray
  varName = ActualVariable
  Set cat = lData.Categories(varName, useShortlabels:=True)
  cat.LowerBound = 1

  'If no categories, just exit
  If cat.Length = 0 Then Exit Sub
  choiceHeader = ValueOf("crf choices")

  'Add headers successively
  If (choiceHeader = "yes") Then
    
    Set cellRng = crfRng.Offset(-2, 1)
    
    For counter = 1 To cat.Length
      cellRng.Value = cat.Item(counter)
      design.ApplyFormat cellRng, scope:=CRFChoiceHeader
      Set cellRng = cellRng.Offset(, 2)
    Next
  End If

  Set cellRng = crfRng.Offset(, 1)

  For counter = 1 To cat.Length
    design.ApplyFormat cellRng, scope:=CRFChoiceValues
    Set cellRng = cellRng.Offset(, 2)
  Next
End Sub


'Build elements on a crf worksheet
Private Sub WriteInfo()

  'For each variable, add the label, the type and the choices
  AddLabel
  AddType
  AddChoices
  
End Sub

Private Sub ICRFVars_WriteInfo()
  WriteInfo
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Checking requirements when building the class
Private Sub CheckRequirements(ByVal varName As String, ll As ILinelist)

    Dim dict As ILLdictionary
    Dim varObject As ILLVariables
    Dim sheetName As String

    If ll Is Nothing Then ThrowError "No linelist object found"
    Set dict = ll.Dictionary()
    Set varObject = LLVariables.Create(dict)
    If Not dict.VariableExists(varName) Then _
     ThrowError "Variable" & varName & "Does not exists!"
    sheetName = varObject.Value(varName:=varName, colName:="sheet name")
    If Not ll.sheetExists(sheetName) Then _
     ThrowError varName & "worksheet named" & sheetName & "is not created"
End Sub