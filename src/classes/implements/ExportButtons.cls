VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ExportButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage export command buttons"

'@Folder("Linelist")
'@ModuleDescription("Manage export command buttons")

Option Explicit
Implements IExportButtons

Private Const CLASSNAME As String = "ExportButtons"
Private Const NAMETAG As String = "CMDExport"

'Objects of the form, the command and the check
Private WithEvents Btn As MSForms.CommandButton
Private chkFilt As MSForms.CheckBox

Private Type TExpBut
  currwb As Workbook
  trads As ITranslation
End Type
Private this As TExpBut

'Exposed methods

Public Function Create(ByVal currwb As Workbook, _
                       ByVal trads As ITranslation, _
                       Byval expBtn As MSForms.CommandButton, _
                       ByVal expChk As MSForms.CheckBox) As IExportButtons

  CheckRequirements currwb, trads, expBtn, expChk

  With New ExportButtons
    Set .CurrentWorkbook = currwb
    Set .MessTrads = trads
    Set .PushButton = expBtn
    Set .CheckButton = expChk
    Set Create = .Self
  End With
End Function


Public Property Set CurrentWorkbook(ByVal currwb As Workbook)
  Set this.currwb = currwb
End Property

Public Property Get CurrentWorkbook() As Workbook
  Set CurrentWorkbook = this.currwb
End Property

Public Property Set MessTrads(ByVal trads As ITranslation)
  Set this.trads = trads
End Property

Public Property Get MessTrads() As ITranslation
  Set MessTrads = this.trads
End Property

Public Property Set PushButton(ByVal expBtn As MSForms.CommandButton)
  Set Btn = expBtn
End Property

Public Property Set CheckButton(ByVal expChk As MSForms.CheckBox)
  Set chkFilt = expChk
End Property

Public Property Get Self() As IExportButtons
    Set Self = Me
End Property

Private Property Get ExportNumber() As Integer
  ExportNumber = CInt(Replace(Btn.Name, NAMETAG, vbNullString))
End Property

Private Property Get UseFilter() As Boolean
  UseFilter = chkFilt.Value
End Property

Private Property Let UseFilter(ByVal expFilt As Boolean)
  chkFilt.Value = expFilt
End Property

'Event Fired when you ask for a typical export
Private Sub Btn_Click()
  CreateExport
End Sub

Private Function TranslatedValue(ByVal msgCode As String) As String
  Dim trads As ITranslation
  Set trads = MessTrads()
  TranslatedValue = trads.TranslatedValue(msgCode)
End Function

Private Sub UpdateData()

  Dim sh As Worksheet                        'The actual worksheet
  Dim filtsh As Worksheet                    'Filtered worksheet
  Dim Lo As ListObject
  Dim destRng As Range
  Dim delRng As Range
  Dim LoRng As Range
  Dim rowCounter As Long
  Dim filtLoHrng As Range 'HeaderRowRange of listObject on filtered sheet
  Dim wsFunc As WorksheetFunction
  Dim wb As Workbook

  On Error GoTo ErrUpdate

  Set wsFunc = Application.WorksheetFunction
  Set wb = CurrentWorkbook()
  
  For Each sh In wb.Worksheets
    If sh.Cells(1, 3).Value = "HList" Then
        'Clean the filtered table list object
        Set Lo = sh.ListObjects(1)
        If Not Lo.DataBodyRange Is Nothing Then
            Set LoRng = Lo.DataBodyRange
            Set filtsh = wb.Worksheets(sh.Cells(1, 5).Value)
            rowCounter = LoRng.Rows.Count
            With filtsh
                On Error Resume Next
                  .ListObjects(1).DataBodyRange.Delete
                On Error GoTo ErrUpdate
                'Lo is the listObject
                'LoRng is the listobject databodyrange in HList
                'destRng is the listObject databodyrange in filtered sheet
            
                .ListObjects(1).Resize .Range(Lo.Range.Address)
                'This is the dataBodyRange of the filtered sheet
                Set destRng = .Range(LoRng.Address)
                Set filtLoHrng = .ListObjects(1).HeaderRowRange
                'Initialize the range to delete at the end of the table
                Set delRng = Nothing
            End With
            'move values to filtered sheet
            destRng.Value = LoRng.Value
            Do While rowCounter >= 1
                If (LoRng.Cells(rowCounter, 1).EntireRow.HIDDEN) Or _ 
                   (wsFunc.CountA(filtLoHrng.Offset(rowCounter)) = 0) Then
                    If delRng Is Nothing Then
                      Set delRng = filtLoHrng.Offset(rowCounter)
                    Else
                      Set delRng = Application.Union(delRng, filtLoHrng.Offset(rowCounter))
                    End If
                End If
                rowCounter = rowCounter - 1
            Loop
            'Delete the range if necessary
            If Not (delRng Is Nothing) Then delRng.Delete
        End If
    End If
  Next
  Exit Sub
  
ErrUpdate:
  MsgBox  TranslatedValue("MSG_ErrUpdate") & ": " & Err.Description, _
          vbCritical + vbOKOnly
End Sub


Private Sub ConfirmFilterUse()    
  Dim confirm As Integer

  If UseFilter() Then
      confirm = MsgBox(TranslatedValue("MSG_AskFilter"), _
                      vbYesNo + vbQuestion, _
                      TranslatedValue("MSG_ThereIsFilter"))
      If confirm = vbYes Then
        'This function is in EventsGlobal Analysis, update filtertables will update all
        'filters in the current workbook.
        UpdateData
      Else
        useFilter = False
      End If
  End If
End Sub

Private Sub CreateExport()

  Dim scope As Integer
  Dim expOut As IOutputSpecs
  Dim trads As ITranslation
  Dim currwb As Workbook

  'Add Error management
  On Error GoTo ErrHand
  BusyApp cursor:=xlWait
  scope = ExportNumber()
  Set currwb = CurrentWorkbook()
  Set trads = MessTrads()
  ConfirmFilterUse
  Set expOut = OutputSpecs.Create(currwb, scope)
  expOut.Save trads, useFilter()
  
  NotBusyApp
  Exit Sub

ErrHand:
  On Error Resume Next
  MsgBox TranslatedValue("MSG_ErrHandExport"), _
          vbOKOnly + vbCritical, _
          TranslatedValue("MSG_Error")
  'Close all oppened workbooks
  expOut.CloseAll
  On Error GoTo 0
  NotBusyApp
End Sub


'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

Private Sub CheckRequirements(ByVal wb As Workbook, ByVal trads As ITranslation, _ 
                              ByVal expBtn As MSForms.CommandButton, _ 
                              ByVal expChk As MSForms.CheckBox)

  If wb Is Nothing Then ThrowError "Unable to initialize the class, provide a workbook"
  If trads Is Nothing Then ThrowError "Unable to initialize the class, provide translation"
  If expBtn Is Nothing Then ThrowError "Unable to initialize the class, Please provide a command button"
  If expChk Is Nothing Then ThrowError "Unable to initialize the class, Please provide export check"
End Sub


Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, errorMessage
End Sub


Private Sub IExportButtons_CreateExport()
  CreateExport
End Sub