VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "LLShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Linelist")
'@ModuleDescription("Linelist Show/Hide Class")

Option Explicit
Implements ILLShowHide

'Exposed methods
Private Const CLASSNAME As String = "LLShowHide"

Private Type TShowHide
  trads As ITranslation
  dict As ILLdictionary
  sh As Worksheet
  SHOW As String
  SHOWN As String
  HIDE As String
  HIDDEN As String
  MANDATORY As String
  DESIGNERHIDDEN As String
  PRINTSHOWHORIZ As String
  PRINTSHOWVERTI As String
  TriggerShowHide As Boolean
  TriggerListClick As Boolean
End Type

Private this As TShowHide
'messages for show/hide and mandatory
Private Const TEMPORARYSHEET As String = "show_hide_temp__"
Private Const PRINTPREFIX As String = "print_"

'Boolean Value to trigger the launch of the show/hide process

Public Function Create(Byval trads As ITranslation, ByVal dict As ILLdictionary, _
                       ByVal sh As Worksheet) As ILLShowHide

  CheckRequirements trads, dict, sh

  With New LLShowHide
    Set .MessagesTrads = trads
    Set .Dictionary = dict
    Set .Wksh = sh
    Set Create = .Self()
  End With
End Function

Public Property Get MessagesTrads() As ITranslation
  Set MessagesTrads = this.trads
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Property Get Dictionary() As ILLdictionary
  Set Dictionary = this.dict
End Property

Public Property Get FormObject() As Object
  Dim scope As ShowHideScope

  scope = ActualScope()
  Select Case scope
  Case ShowHideScopeLinelist
    Set FormObject = [F_ShowHideLL]
  Case ShowHideScopePrint
    Set FormObject = [F_ShowHidePrint]
  End Select
End Property

Public Property Set MessagesTrads(ByVal trads As ITranslation)
  Set this.trads = trads
End Property

Public Property Get ActualScope()
  'Actual scope is in the worksheet
  Dim sh As Worksheet
  Dim rngVal As String
  Dim scope As ShowHideScope

  Set sh = Wksh()
  On Error Resume Next
    rngVal = sh.Cells(1, 3).Value
  On Error GoTo 0
  If rngVal = "HList" Then
    scope = ShowHideScopeLinelist
  ElseIf rngVal = "HList Print" Then
    scope = ShowHideScopePrint
  End If
  ActualScope = scope
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
  Set this.dict = dict
End Property

Public Property Get Self() As ILLShowHide
  Set Self = Me
End Property

Private Sub InitializeMessages(ByVal trads As ITranslation)
  this.SHOW = trads.TranslatedValue("MSG_Show")
  this.SHOWN = trads.TranslatedValue("MSG_Shown")
  this.HIDE = trads.TranslatedValue("MSG_Hide")
  this.HIDDEN = trads.TranslatedValue("MSG_Hidden")
  this.MANDATORY = trads.TranslatedValue("MSG_Mandatory")
  this.DESIGNERHIDDEN = trads.TranslatedValue("MSG_DesHidden")
  this.PRINTSHOWHORIZ = trads.TranslatedValue("MSG_ShowHoriz")
  this.PRINTSHOWVERTI = trads.TranslatedValue("MSG_ShowVerti")
End Sub

Private Function TempSheet() As Worksheet
  Dim wb As Workbook
  Dim sh As Worksheet
  Dim tpsh As Worksheet

  Set sh = Wksh()
  Set wb = sh.Parent
  Set TempSheet = wb.Worksheets(TEMPORARYSHEET)
End Function

'Subs to Enable/disable the show Hide process
Private Sub EnableShowHide()
  this.TriggerShowHide = True
End Sub

Private Sub DisableShowHide()
  this.TriggerShowHide = False
End Sub

Private Function InShowHideMode() As Boolean
  InShowHideMode = this.TriggerShowHide
End Function

Private Function InClickMode() As Boolean
  InClickMode = this.TriggerListClick
End Function

Private Sub EnableClickMode()
  this.TriggerListClick = True
End Sub

Private Sub DisableClickMode()
  this.TriggerListClick = False
End Sub

Private Function ShowHideTable() As Range

  Const SHOWHIDEPREF As String = "ShowHideTable_"

  Dim tpsh As Worksheet
  Dim Lo As ListObject
  Dim loName As String
  Dim dict As ILLdictionary
  Dim retRng As Range
  Dim vars As ILLvariables
  Dim sh As Worksheet
  Dim varRng As Range
  Dim showHideColumn As String
  Dim sheetName As String
  Dim cellRng As Range
  Dim tempCounter As Long
  Dim scope As ShowHideScope
  Dim counterSheetName As String
  Dim mainLab As String
  Dim controlVal As String
  Dim varName As String
  Dim loRng As Range
  Dim startCol As Long
  Dim tabName As String
  Dim defaultValue As String
  Dim statusValue As String

  Set tpsh = TempSheet()
  Set sh = Wksh()
  tabName = sh.Cells(1, 4).Value
  loName = SHOWHIDEPREF & tabName

  
  With tpsh
    If Not LoExists(loName) Then

      'Create the listObject
      Set dict = Dictionary()
      Set varRng = dict.DataRange("variable name")
      Set vars = LLVariables.Create(dict)
      scope = ActualScope()
      showHideColumn = IIf(scope = ShowHideScopePrint, "register book", "visibility")
      defaultValue = IIf(scope = ShowHideScopePrint, this.PRINTSHOWHORIZ, this.SHOWN)
      'Get the sheet name
      sheetName = IIf(scope = ShowHideScopePrint, _
                      Replace(sh.Name, PRINTPREFIX, vbNullString), _
                      sh.Name)
      sheetName = Application.WorksheetFunction.Trim(sheetName)
      Set cellRng = varRng.Cells(1, 1)

      'Define the start column
        startCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
        tempCounter = 2
        'Loop for all variables in the dictionary (I don't presume the dictionary is sorted on sheet names)
        Do While cellRng.Value <> vbNullString
          varName = cellRng.Value
          counterSheetName = vars.Value(varName:=varName, colName:="sheet name")
          If counterSheetName = sheetName Then
            'If it is not hidden by the designer
            If (vars.Value(varName:=varName, colName:="status") <> "hidden") Then
              'main label
              mainLab =  vars.Value(varName:=varName, colName:="main label")
              controlVal = vars.Value(varName:=varName, colName:="control")

              'For geo variables, add the main label and the section
              If (InStr(1, controlVal, "geo") > 0) Or (InStr(1,controlVal, "hf") > 0) _
              Then mainLab = vars.Value(varName:=varName, colName:="sub section") & ": " & mainLab

              .Cells(tempCounter, startCol + 1).Value = mainLab

              'variable name
              .Cells(tempCounter, startCol + 2).Value = varName

              'Status
              statusValue = vars.Value(varName:=varName, colName:=showHideColumn)

              Select Case statusValue
              Case "mandatory"
                statusValue = this.MANDATORY
              Case "optional"
                statusValue = this.SHOWN
              Case "print, vertical header"
                statusValue = this.PRINTSHOWVERTI
              Case "print, horizontal header"
                statusValue = this.PRINTSHOWHORIZ
              Case "hidden"
                statusValue = this.HIDDEN
              End Select

              'On admin 1 and Linelist, the status is mandatory
              If (controlVal = "geo1") And (scope = ShowHideScopeLinelist) Then _
               statusValue = this.MANDATORY

              'The default value is the value to write if nothing found as status
              '(empty registerbook column)             

              If statusValue = vbNullString Then statusValue = defaultValue

              .Cells(tempCounter, startCol + 3).Value = statusValue
              'column index (needed for further usages)
              .Cells(tempCounter, startCol + 4).Value = vars.Value(varName:=varName, colName:="column index")
              'Control
              .Cells(tempCounter, startCol + 5).Value = vars.Value(varName:=varName, colName:="control")
              tempCounter = tempCounter + 1
            End If
          End If
          Set cellRng = cellRng.Offset(1)
        Loop
        .Cells(1, startCol + 1).Value = "main label"
        .Cells(1, startCol + 2).Value = "variable name"
        .Cells(1, startCol + 3).Value = "status"
        .Cells(1, startCol + 4).Value = "column index"
        .Cells(1, startCol + 5).Value = "control"
        'substract 1 for resizing purposes
        tempCounter = tempCounter - 1
        Set loRng = .Range(.Cells(1, startCol + 1), .Cells(tempCounter, startCol + 5))
        .ListObjects.Add(xlSrcRange, loRng, , xlYes).Name = loName
    End If
    Set Lo = .ListObjects(loName)
    
    If Not (Lo.DataBodyRange Is Nothing) Then
      Set retRng = Lo.DataBodyRange
      Set retRng = .Range(retRng.Cells(1, 1), retRng.Cells(retRng.Rows.Count, 3))
    Else
      Set retRng = Lo.Range
    End If
  End With

  Set ShowHideTable = retRng
End Function

Private Function Value(ByVal lineNumber As Long, Optional ByVal colName As String = "status") As String
  Dim formRng As Range
  Dim colIndex As Long

  Set formRng = ShowHideTable()

  Select Case colName
  Case "status"
    colIndex = 3
  Case "main label"
    colIndex = 1
  Case "column index"
    colIndex = 4
  Case "variable name"
    colIndex = 2
  Case Else
    colIndex = 3
  End Select

  On Error Resume Next
  Value = formRng.Cells(lineNumber, colIndex).Value
  On Error GoTo 0
End Function

'Change the status in the table
Private Sub SetStatus(ByVal lineNumber As Long, ByVal newStatus As String)
  Dim formRng As Range
  Dim colIndex As Long

  Set formRng = ShowHideTable()
  'Change the corresponding value in the temporary sheet
  formRng.Cells(lineNumber, 3).Value = newStatus
End Sub


Private Sub Load(ByVal tradsform As ITranslation)
  Dim formRng As Range 'Range to update the form
  Dim tabName As String 'Table name of the worksheet
  Dim fobj As Object
  Dim trads As ITranslation
  Dim scope As ShowHideScope
  Dim listObject As Object

  Set trads = MessagesTrads()

  'Initialize translations
  InitializeMessages trads

  'Range of the form
  Set formRng = ShowHideTable()
  Set fobj = FormObject()
  scope = ActualScope()

  'Initialize the values
  If scope = ShowHideScopeLinelist Then
    Set listObject =  fobj.LST_LLVarNames
  Else
    Set listObject = fobj.LST_PrintNames
  End If

  listObject.BoundColumn = 2
  listObject.ColumnCount = 3
  listObject.List = formRng.Value

  DisableClickMode

  listObject.Selected(0) = True

  EnableClickMode
  EnableShowHide

  'Load the forms
  fobj.Caption = tradsform.TranslatedValue(fobj.Name)
  tradsform.TranslateForm fobj

  'Add caption to option buttons
  On Error Resume Next
    fobj.OPT_Show.Caption = this.SHOW
    fobj.OPT_PrintShowHoriz.Caption = this.PRINTSHOWHORIZ
    fobj.OPT_PrintShowVerti.Caption = this.PRINTSHOWVERTI
    fobj.OPT_Hide.Caption = this.HIDE
  On Error GoTo 0

  fobj.Show
End Sub

'Update visibility when you select one line of the show hide
Private Sub UpdateVisibility(ByVal indexValue As Long)

  Dim scope As ShowHideScope
  Dim fobj As object

  If Not InClickMode Then Exit Sub

  'disable the show/hide mode
  DisableShowHide

  Set fobj = FormObject()
  scope = ActualScope()

  Select Case Value(indexValue, "status")

  Case this.PRINTSHOWHORIZ

    fobj.OPT_PrintShowHoriz.Caption = this.PRINTSHOWHORIZ
    fobj.OPT_PrintShowHoriz.Value = 1

  Case this.PRINTSHOWVERTI

    fobj.OPT_PrintShowVerti.Caption = this.PRINTSHOWVERTI
    fobj.OPT_PrintShowVerti.Value = 1

  'This is the scope linelist (show)
  Case this.SHOWN

    fobj.OPT_Show.Value = 1
    fobj.OPT_Show.Caption = this.SHOW
    fobj.OPT_Show.Top = 75

    fobj.OPT_Hide.Visible = True
    fobj.OPT_Show.Visible = True
    fobj.OPT_Hide.Value = 0

  Case this.MANDATORY

    fobj.OPT_Show.Caption = this.MANDATORY
    fobj.OPT_Show.Top = 100
    fobj.OPT_Hide.Visible = False

  'This is both for scope linelist and print sheet
  Case this.HIDDEN

    fobj.OPT_Hide.Visible = True
    fobj.OPT_Hide.Value = 1
    
    If scope = ShowHideScopeLinelist Then
      fobj.OPT_Show.Caption = this.SHOW
      fobj.OPT_Show.Top = 75
    End If
  Case Else 'Show Everything, and select show horizontally on printed sheet
    On Error Resume Next
    fobj.OPT_Show.Value = 1
    fobj.OPT_PrintShowHoriz.Value = 1
    fobj.OPT_Hide.Visible = True
    fobj.OPT_Show.Caption = this.SHOW
    On Error GoTo 0
  End Select

  'Enable the Show/Hide process
  EnableShowHide
End Sub

Private Sub ShowHideLogic(ByVal indexValue As Long)



  Const TABPRINTPREFIX As String = "pr" 'Print prefix for tables

  Dim sh As Worksheet
  Dim fobj As Object
  Dim scope As ShowHideScope
  Dim colIndex As Long
  Dim cellRng As Range
  Dim orientationValue As Long
  Dim tabName As String
  Dim showValue As String
  Dim formRng As Range
  Dim listObject As Object
  Dim hideCol As Boolean

  On Error GoTo ErrHand:

  Set sh = Wksh()
  Set fobj = FormObject()
  scope = ActualScope()
  tabName = sh.Cells(1, 4).Value
  Set formRng = ShowHideTable() 'To return back showhide table afterwards

  If Not InShowHideMode() Or indexValue < 0 Then Exit Sub
  'For mandatory, exit
  If Value(indexValue, "status") = this.MANDATORY Then Exit Sub
  'Get the column index of the variable in the actual worksheet
  On Error Resume Next
    colIndex = -1
    colIndex = CLng(Value(indexValue, "column index"))
  On Error GoTo 0
  If colIndex = -1 Then Exit Sub

  'listObject depending of the scope
  If scope = ShowHideScopeLinelist Then
    'Linelist
    Set listObject = fobj.LST_LLVarNames

    If fobj.OPT_Hide.Value Then
      hideCol = True
      showValue = this.HIDDEN
    ElseIf fobj.OPT_Show.Value Then
      hideCol = False
      showValue = this.SHOWN
    End If

  Else
    'Linelist print
    Set listObject = fobj.LST_PrintNames

    On Error Resume Next
      Set cellRng = sh.Range(Replace(tabName, TABPRINTPREFIX, vbNullString) & _
                             "_" & "PRINTSTART")
    On Error GoTo 0
    'Unable to find the print start Range, exit
    If cellRng Is Nothing Then Exit Sub
    'This is the main label header of the variable
    Set cellRng = sh.Cells(cellRng.Offset(-2).Row, colIndex)

    If fobj.OPT_Hide.Value Then
      hideCol = True
      showValue = this.HIDDEN
    ElseIf fobj.OPT_PrintShowHoriz.Value Then
      hideCol = False
      showValue = this.PRINTSHOWHORIZ
      orientationValue = 0
    ElseIf fobj.OPT_PrintShowVerti.Value Then
      hideCol = False
      showValue = this.PRINTSHOWVERTI
      orientationValue = 90
    End If
    'unhide and change the orientation
    cellRng.EntireRow.RowHeight = 100
    cellRng.Orientation = orientationValue
    cellRng.EntireColumn.AutoFit
  End If

  sh.Columns(colIndex).Hidden = hideCol
  SetStatus indexValue, showValue
  listObject.Clear
  listObject.List = formRng.Value
  DisableClickMode
  listObject.Selected(indexValue - 1) = True
  EnableClickMode

ErrHand:

End Sub

Private Sub ChangeColWidth(ByVal indexValue As Long)

  Dim sh As Worksheet
  Dim colIndex As Long
  Dim trads As ITranslation
  Dim inputValue As String
  Dim actualWidth As Long

  If ActualScope = ShowHideScopeLinelist Then Exit Sub

  On Error Resume Next
    colIndex = -1
    colIndex = CLng(Value(indexValue, "column index"))
  On Error GoTo 0
  If colIndex = -1 Then Exit Sub 'Find the column index

  Set sh = Wksh()
  Set trads = MessagesTrads()

  Do While (True)
    inputValue = InputBox(trads.TranslatedValue("MSG_ColWidth"), trads.TranslatedValue("MSG_Enter"))
    If (inputValue = vbNullString) Then Exit Sub
    If IsNumeric(inputValue) Then Exit Do
    If (MsgBox (trads.TranslatedValue("MSG_EnterNumeric"), vbOkCancel, "") = vbCancel) Then Exit Sub
  Loop

  On Error Resume Next
    actualWidth = CLng(inputValue)
    sh.Columns(colIndex).ColumnWidth = actualWidth
  On Error GoTo 0
End Sub

'Errors and Requirements
Private Sub CheckRequirements(ByVal trads As ITranslation, ByVal dict As ILLdictionary, ByVal sh As Worksheet)

  Dim wb As Workbook
  Dim tpsh As Worksheet

  'Translations
  If trads Is Nothing Then ThrowError "Error, no translation found"
  'Dictionary
  If dict Is Nothing Then ThrowError "Error, no dictionary found"
  If Not dict.Prepared Then ThrowError "Error, the dictionary should be prepared"
  If sh Is Nothing Then ThrowError "Error, unable to find the worksheet where to hide/show variables"
  Set wb = sh.Parent
  On Error Resume Next
  Set tpsh = wb.Worksheets(TEMPORARYSHEET)
  On Error GoTo 0
  If tpsh Is Nothing Then ThrowError "Error, no temporary sheet find in current workbook (required for computations)"
End Sub


Private Function LoExists(ByVal loName As String) As Boolean

  Dim Lo As ListObject
  Dim sh As Worksheet
  Set sh = TempSheet()
  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0
  LoExists = Not (Lo Is Nothing)
End Function

'Throw Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub ILLShowHide_Load(ByVal tradsform As ITranslation)
   Load tradsform
End Sub

Private Sub ILLShowHide_UpdateVisibilityStatus(ByVal index As Long)
   UpdateVisibility indexValue:=(index + 1)
End Sub

Private Sub ILLShowHide_ShowHideLogic(ByVal index As Long)
 ShowHideLogic indexValue:=(index + 1)
End Sub

Private Sub ILLShowHide_ChangeColWidth(ByVal index As Long)
  ChangeColWidth indexValue:=(index + 1)
End Sub
