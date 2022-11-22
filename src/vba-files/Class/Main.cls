VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("designer")
'@ModuleDescription("main interface to linelist specifications paths")
'@IgnoreModule

Option Explicit
Implements IMain

Private Type TMain
  sh As Worksheet
End Type

Private this As TMain

Public Function Create(sh As Worksheet) As IMain

  CheckRequirements sh

  With New Main
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

Public Property Get Self() As IMain
  Set Self = Me
End Property

Private Property Get OutputPath() As String
  Dim sh As Worksheet
  Set sh = Wksh()
  OutputPath = sh.Range("RNG_LLDir").Value
End Property

Private Property Get LinelistName() As String
  Dim sh As Worksheet
  Set sh = Wksh()
  LinelistName = sh.Range("RNG_LLName").Value
End Property

Private Property Get SetupPath() As String
  Dim sh As Worksheet
  Set sh = Wksh()
  SetupPath = sh.Range("RNG_PathDico").Value
End Property

Private Property Get GeoPath() As String
  Dim sh As Worksheet
  Set sh = Wksh()
  GeoPath = sh.Range("RNG_PathGeo").Value
End Property

Private Property Get Ready() As Boolean
  Dim sh As Worksheet
  Set sh = Wksh()
  Ready = (sh.Range("RNG_MainReady").Value = "ready")
End Property


'Check if a workbook is Opened
Public Function WorkbookIsOpen(wkbName As String) As Boolean
  Dim wb As Workbook                         'Just try to set the workbook if it fails it is closed
  On Error Resume Next
  Set wb = Application.Workbooks.Item(wkbName)
  WorkbookIsOpen = (Not wb Is Nothing)
  On Error GoTo 0
End Function


'Check if the main sheet is ready for generation process

Public Sub CheckReadiness(ByVal Trads As IDesTranslation)
  Dim sh As Worksheet
  Dim wkbName As String 'Eventual Workbook name to check if a workbook is opened
  Dim redColor As Variant
  Dim whiteColor As Variant

  Set sh = Wksh()
  redColor = RGB(252, 228, 214)
  whiteColor = vbWhite

  sh.Range("RNG_MainReady").Value = "not ready"
  sh.Range("RNG_MainReady").Font.color = redColor

  'Be sure the dictionary path is not empty
  If sh.Range("RNG_PathDico").Value = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathDic")
      sh.Range("RNG_PathDico").Interior.color = redColor
      Exit Sub
  End If

  'Now check if the file exists
  wkbName = Dir(sh.Range("RNG_PathDico").Value)

  If wkbName = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathDic")
      sh.Range("RNG_PathDico").Interior.color = redColor
      Exit Sub
  End If

  'Be sure the dictionnary is not opened
  If WorkbookIsOpen(wkbName) Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_CloseDic")
      sh.Range("RNG_PathDico").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_AlreadyOpen"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_Title_Dictionnary")
      Exit Sub
  End If

  sh.Range("RNG_PathDico").Interior.color = whiteColor 'if path is OK

  'Checking coherence of the GEO  ------------------------------------------------

  'Be sure the geo path is not empty
  If sh.Range("RNG_PathGeo").Value = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathGeo")
      sh.Range("RNG_PathGeo").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_PathGeo"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_TitleGeo")
      Exit Sub
  End If

  'Now check if the file exists
  If Dir(sh.Range("RNG_PathGeo").Value) = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathGeo")
      sh.Range("RNG_PathGeo").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_PathGeo"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_TitleGeo")
      Exit Sub
  End If

  sh.Range("RNG_PathGeo").Interior.color = whiteColor 'if path is OK

  'Checking coherence of the Linelist File ------------------------------------------------------

  'Be sure the linelist directory is not empty
  If sh.Range("RNG_LLDir").Value = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathLL")
      sh.Range("RNG_LLDir").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_PathLL"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_TitleLL")
      Exit Sub
  End If

  'Be sure the linelist is not opened
  wkbName = Dir(sh.Range("RNG_LLDir").Value & Application.PathSeparator & sh.Range("RNG_LLName").Value & ".xlsb")
  
  If WorkbookIsOpen(wkbName) Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_CloseOutPut")
      sh.Range("RNG_LLDir").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_CloseOutPut"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_Title_OutPut")
      Exit Sub
  End If

  'Be sure the directory for the linelist exists
  If Dir(sh.Range("RNG_LLDir").Value & "*", vbDirectory) = vbNullString Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_PathLL")
      sh.Range("RNG_LLDir").Interior.color = redColor
      MsgBox Trads.TranslationMsg("MSG_PathLL"), vbExclamation + vbOKOnly, Trads.TranslationMsg("MSG_TitleLL")
      Exit Sub
  End If

  sh.Range("RNG_LLDir").Interior.color = whiteColor 'if path is OK

  'Checking coherence of the linelist name ------------------------------------------------------

  'be sure the linelist name is not empty
  If sh.Range("RNG_LLName") = "" Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_LLName")
      sh.Range("RNG_LLName").Interior.color = redColor
      Exit Sub
  End If

  'Be sure the linelist workbook is not already opened
  If WorkbookIsOpen(sh.Range("RNG_LLName").Value & ".xlsb") Then
      sh.Range("RNG_Edition").Value = Trads.TranslationMsg("MSG_CloseLL")
      sh.Range("RNG_LLName").Interior.color = redColor
      Exit Sub
  End If

  'Be sure the linelist name is well written (or change it)
  sh.Range("RNG_LLName").Value = CorrectLinelistName(sh.Range("RNG_LLName").Value)

  'If everything is ready
  sh.Range("RNG_MainReady").Value = "ready"
  sh.Range("RNG_MainReady").Font.color = whiteColor
End Sub


Private Function CorrectLinelistName(ByVal llfileName As String) As String
    
  'In the file name, replace forbidden characters with an underscore
  CorrectLinelistName = vbNullString
  Dim FileName As String
  FileName = Replace(llfileName, "<", "_")
  FileName = Replace(FileName, ">", "_")
  FileName = Replace(FileName, ":", "_")
  FileName = Replace(FileName, "|", "_")
  FileName = Replace(FileName, "?", "_")
  FileName = Replace(FileName, "/", "_")
  FileName = Replace(FileName, "\", "_")
  FileName = Replace(FileName, "*", "_")
  FileName = Replace(FileName, ".", "_")
  FileName = Replace(FileName, """", "_")
  
  CorrectLinelistName = Application.WorksheetFunction.Trim(FileName)
End Function

'Update the status Bar in the main sheet
Private Sub UpdateStatus(ByVal numValue As Integer)

  Const NUMBEROFBARS As Integer = 100
  
  Dim CurrentStatus As Integer
  Dim currentEvent As Boolean
  Dim sh As Worksheet

  Set sh = Wksh()

  currentEvent = Application.ScreenUpdating
  Application.ScreenUpdating = True
  CurrentStatus = (NUMBEROFBARS) * Round(numValue / 100, 1)
  sh.Range("RNG_Update").Value = "[" & String(CurrentStatus, "|") & Space(NUMBEROFBARS - CurrentStatus) & "]" & _
                                  " " & CInt(numValue) & "%"
  Application.ScreenUpdating = currentEvent
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub CheckRequirements(ByVal sh As Worksheet)

  Dim rngTable As BetterArray
  Dim counter As Long
  Dim rngName As String

  'Sheet Should Exists
  If sh Is Nothing Then ThrowError "Geo sheet not defined"
  Set rngTable = New BetterArray
  rngTable.Push "RNG_LangDesigner", "RNG_LLForm", "RNG_LangSetup", "RNG_DesignLL", _
                "RNG_PathDico", "RNG_PathGeo", "RNG_LLDir", "RNG_LLName", _
                "RNG_Edition", "RNG_MainReady", "RNG_Update"
  
  For counter = rngTable.LowerBound To rngTable.UpperBound
    rngName = rngTable.Item(counter)
    RangeExists sh, rngName
  Next
End Sub


Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)

  Dim rng As Range
  On Error Resume Next
  Set rng = sh.Range(rngName)
  On Error GoTo 0

  If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name
End Sub

Public Property Get Imain_OutputPath() As String
  Imain_OutputPath = OutputPath()
End Property

Public Property Get Imain_LinelistName() As String
  Imain_LinelistName = LinelistName()
End Property

Public Property Get Imain_SetupPath() As String
  Imain_SetupPath = SetupPath()
End Property

Public Property Get Imain_GeoPath() As String
  Imain_GeoPath = GeoPath()
End Property

Public Property Get Imain_Ready() As Boolean
  Imain_Ready = Ready()
End Property

Public Sub Imain_CheckReadiness(ByVal Trads As IDesTranslation)
  CheckReadiness Trads
End Sub

Public Sub Imain_UpdateStatus(ByVal numValue As Integer)
  UpdateStatus numValue
End Sub
