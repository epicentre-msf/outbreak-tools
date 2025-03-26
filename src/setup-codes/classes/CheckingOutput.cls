VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "How to write the output of all checks"

'@PredeclaredId
'@Folder("Checks")
'@ModuleDescription("How to write the output of all checks")

Option Explicit
Implements ICheckingOutput

Private Type TCheckOut
  sh As Worksheet
End Type

Private this As TCheckOut
Private Const CLASSNAME As String = "CheckingOutput"


'Exposed methods
Public Function Create(ByVal sh As Worksheet) As ICheckingOutput

  If sh Is Nothing Then _
  ThrowError "Unable to create an instance of the class, please provide a valid worksheet"

  With New CheckingOutput
    Set .Wksh = sh
    Set Create = .Self()
  End With
End Function

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Property Get Self() As ICheckingOutput
  Set Self = Me
End Property

'Find the CellRange where to start writing
Private Property Get WritingCell() As Range
  Dim sh As Worksheet
  Set sh = Wksh()
  With sh
    Set WritingCell = .Cells(.Cells(.Rows.Count, 2).End(xlUp).Row, 2).Offset(3)
  End With
End Property

'You can write either at level = 1 for title, level = 2 for subtitle
'and level = 3 for simple text
Private Sub WriteValue(ByVal level As Byte, textValue As String, _
                       Optional ByVal color As String = "black")
  Dim fontCol As Long
  Dim intCol As Long
  Dim sh As Worksheet
  Dim cellRng As Range
  Dim counter As Integer

  Set sh = Wksh()
  Set cellRng = WritingCell()
  If level = 2 Then Set cellRng = cellRng.Offset(-1)
  If level = 3 Then Set cellRng = cellRng.Offset(-2)
  'Set the color of the text
  Select Case color
    Case "blue"
      fontCol = RGB(47, 117, 181)
      intCol = RGB(217, 225, 242)
    Case "red"
      fontCol = RGB(255, 0, 0)
      intCol = RGB(252, 228, 214)
    Case "orange"
      fontCol = RGB(128, 96, 0)
      intCol = RGB(255, 242, 204)
    Case "grey"
      fontCol = RGB(160, 160, 160)
      intCol = RGB(250, 250, 250)
    Case Else
      fontCol = vbBlack
      intCol = RGB(208, 206, 206)
  End Select

  On Error Resume Next
    'If there is no -- The split can be just one character (not important though)
    cellRng.Cells(1, 1).Value = Split(textValue, "--")(0)
    cellRng.Cells(1, 2).Value = Split(textValue, "--")(1)
    cellRng.Cells(1, 3).Value = Split(textValue, "--")(2)
  On Error GoTo 0

  Set cellRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(1, 3))

  With cellRng
    .Font.Size = Switch(level = 1, 14, level = 2, 12, level = 3, 10)
    .Interior.color = intCol
    .Font.color = fontCol
    .Font.Bold = False
    .EntireColumn.AutoFit
    On Error Resume Next
    .Font.Name = "Tahoma"
    On Error GoTo 0
    For counter = 7 To 12
        With .Borders(counter)
            .LineStyle = xlDashDotDot
            .Weight = xlHairline
            .color = RGB(10, 15, 100)
        End With
    Next
  End With
End Sub

'Write One checking
Private Sub WriteChecking(ByVal check As IChecking)
  Dim counter As Long
  Dim keysList As BetterArray
  Dim keyName As String
  Dim title As String
  Dim subTitle As String
  Dim sh As Worksheet
  Dim keyColor As String
  Dim keyType As String
  Dim keyLabel As String

  Set keysList = check.ListOfKeys()
  Set sh = Wksh()
  title = check.Heading()
  subTitle = check.Heading(subTitle:=True)

  'Write the title if needed
  If (sh.Cells.Find(What:=Split(title, "--")(0), lookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
    WriteValue level:=1, textValue:=title
  End If

  'write the subtitle if needed
  If subTitle <> vbNullString Then WriteValue level:=2, textValue:=subTitle

  For counter = keysList.LowerBound To keysList.UpperBound
    keyName = keysList.Item(counter)
    keyColor = check.ValueOf(keyName:=keyName, retrievedValue:=checkingColor)
    keyLabel = check.ValueOf(keyName:=keyName, retrievedValue:=checkingLabel)
    keyType = check.ValueOf(keyName:=keyName, retrievedValue:=checkingType)
    keyLabel = keyType & "--" & keyLabel
    WriteValue level:=3, textValue:=keyLabel, color:=keyColor
  Next
End Sub



'Write multiple checkings
'This will write a list of multiple checkings
Private Sub WriteCheckingTables(ByVal checkTable As BetterArray)
  Dim check As IChecking
  Dim counter As Long
  Dim sh As Worksheet

  Set sh = Wksh()
  sh.Cells.Clear
  sh.Cells(1, 2).Value = "Show only:"

  For counter = checkTable.LowerBound To checkTable.UpperBound
    If TypeName(checkTable.Item(counter)) <> "Checking" Then _
    ThrowError _
    "You should provide a betterarray of checkings, item " & _
     counter & "is not of type IChecking"

    Set check = checkTable.Item(counter)

    'Write the check
    WriteChecking check
  Next
End Sub

'Errors and output
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub


'Output elements
Private Sub ICheckingOutput_PrintOutput(ByVal checkTable As BetterArray)
  WriteCheckingTables checkTable
End Sub

