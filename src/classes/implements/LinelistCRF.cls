VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "LinelistCRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Dictionary")
'@ModuleDescription("Designing and Working on the CRF worksheet")
'@IgnoreModule  UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILinelistCRF

Private Type TLLCRF
  ll As ILinelist
  shName As String
End Type

Private this As TLLCRF
Private const CLASSNAME As String = "LinelistCRF"


Public Function Create(ByVal sheetName As String, ll As ILinelist) As ILinelistCRF
  
  CheckRequirements sheetName, ll

  With New LinelistCRF
    .Name = sheetName
    Set .LinelistObject = ll
    Set Create = .Self
  End With

End Function

Public Property Get Self() As ILinelistCRF
  Set Self = Me
End Property

Public Property Let Name(ByVal sheetName As String)
  this.shName = sheetName
End Property

Public Property Get Name() As String
  Name = this.shName
End Property

Public Property Get LinelistObject() As ILinelist
  Set LinelistObject = this.ll
End Property

Public Property Set LinelistObject(ByVal ll As ILinelist)
  Set this.ll = ll
End Property

Private Property Get Wksh() As Worksheet
  Dim ll As ILinelist
  Set Wksh = ll.CRFWksh(Name)
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

'Busy app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'variables object
Private Property Get Vars() As ILLVariables
    Dim dict As ILLdictionary
    Set dict = Dictionary()
    Set Vars = LLVariables.Create(dict)
End Property


'Create the crf worksheet and hide it if
'required
Private Sub Initialize()

  Dim ll As ILinelist
  Dim sh As Worksheet
  Dim design As ILlFormat
  Dim sheetName As String

  Set ll = LinelistObject()
  Set design = DesignFormat()
  sheetName = Name()
  
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
  design.ApplyFormat sh.Cells(1, 3), scope:=LinelistHiddenCell
  design.ApplyFormat sh.Cells(1, 4), scope:=LinelistHiddenCell
End Sub


Private Property Get VarCRFRange(ByVal varName As String) As Range

  Dim sh As Worksheet
  Dim crfLine As Long
  Dim varobj As ILLVariables
  
  Set varobj = Vars()
  Set sh = Wksh()
  
  On Error Resume Next
  crfLine = varobj.Value(colName:="crf index", varName:=varName)
  On Error GoTo 0

  If (crfLine = 0) Then 
    ThrowError "Error when creating CRF Range: Unable to find the crf index for " & varName
  End If

  Set VarCRFRange = sh.Cells(crfLine, 1)
End Property

'Add Label on a variable

Private Sub AddLabel(ByVal varName As String)

  Dim varobj As ILLVariables
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

  Set varobj = Vars()
  Set sh = Wksh()
  Set design = DesignFormat()

  varMainLab = varobj.Value(colName:="main label", varName:=varName)
  varSubLab = varobj.Value(colName:="sub label", varName:=varName)
  varControl = varobj.Value(colName:="control", varName:=varName)
  varStatus = varobj.Value(colName:="status", varName:=varName)

  'Make sure the line exists
  On Error Resume Next
  crfLine = CLng(varobj.Value(colName:="crf index", varName:=varName))
  On Error GoTo 0
  If crfLine = 0 Then Exit Sub
  
  'Get the crf range
  Set crfRng = sh.Cells(crfLine, 1)

  'On the first cell of a crf worksheet, add an anchor
  If crfLine = 5 Then
    tabName = varobj.Value(colName:="table name", varName:=varName)
    sh.Cells(1, 4).Value = tabName
    crfRng.Name = tabName & "_CRFSTART"
  End If

  mainLabLength = Len(varMainLab)

  If (varSubLab <> vbNullStriing) Then
    varMainLab = varMainLab & chr(10) & varSubLab
  End If

  'writing the main label and formatting
  crfRng.Value = varMainLab
  design.ApplyFormat crfRng.Offset(-2).Characters(Start:=1, Length:=mainLabLength), scope:=HListMainLab                      
  design.ApplyFormat crfRng.Characters(Start:=mainLabLength + 1), scope:=HListSublab

  'Hidding the columns not to put in crfs
  hiddenRow = (varStatus = "hidden" Or varStatus = "optional, hidden") Or _ 
              (varControl = "formula" Or varControl = "choice_formula" Or varControl = "case_when")

  crfRng.EntireRow.HIDDEN = hiddenRow
End Sub


Private Sub AddType(ByVal varName As String)

  Dim varobj As ILLVariables
  Dim sh As Worksheet
  Dim crfRng As Range
  Dim varTypes As String
  Dim varFormat As String
  Dim varControl As String
  Dim crfRng As Range
  Dim scope As FormatScopes
  Dim nbDecimals As Integer
  Dim design As ILLFormat


  Set varobj = Vars()
  Set sh = Wksh()
  Set design = DesignFormat()
  Set crfRng = VarCRFRange(varName)

  varTypes = LCase(varobj.Value(colName:="variable type", varName:=varName))
  varFormat = LCase(varobj.Value(colName:="variable format", varName:=varName))
  varFormat = Application.WorksheetFunction.Trim(varFormat)
  varControl = varobj.Value(colName:="control", varName:=varName)
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


Private Sub AddChoices(ByVal varName As String)

  Dim varobj As ILLVariables
  Dim choiceHeader As String
  Dim lData As ILinelistSpecs
  Dim cat As BetterArray
  Dim crfRng As Range
  Dim cellRng As Range
  Dim counter As Long

  Set varobj = Vars()
  varControl = varobj.Value(colName:="control", varName:=varName)
  
  If varControl <> "choice_manual" Then Exit Sub
  
  Set lData = LinelistData()
  Set crfRng = VarCRFRange(varName)
  Set cat = New BetterArray
  Set cat = lData.Categories(varName)
  cat.LowerBound = 1

  'If no categories, just exit
  If cat.Length = 0 Then Exit Sub
  choiceHeader = varobj.Value(colName:="crf choices", varName:=varName)

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
Private Sub Build()

  Dim dict As ILLdictionary
  Dim varRng As Range
  Dim counter As Long
  Dim varName As String


  'Initialize the worksheet
  Initialize

  'For each variable, add the label, the type and the choices
  Set dict = Dictionary()
  Set varRng = dict.DataRange("variable name")

  For counter = 1 To varRng.Rows.Count
    varName = varRng.Cells(counter, 1).Value
    addLabel varName:=varName
    addType varName:=varName
    addChoices varName:=varName
  Next

End Sub


Private Sub ILinelistCRF_Build()
  Build
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Checking requirements when building the class
Private Sub CheckRequirements(ByVal sheetName As String, ByVal ll As ILinelist)
    Dim llshs As ILLSheets
    Dim dict As ILLdictionary

    If ll Is Nothing Then ThrowError "Please provide a linelist object"
    Set dict = ll.Dictionary()
    Set llshs = LLSheets.Create(dict)
    If Not llshs.Contains(sheetName) Then ThrowError "The provided sheetname " & sheetName _
                                   & " is not present in the dictionary"
End Sub