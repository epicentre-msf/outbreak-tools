VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ExportButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit

Private WithEvents btn As MSForms.CommandButton

Private Type TExpBut
  currwb As Workbook
  tradsmess As ITranslation
  fObj As Object
End Type


'Exposed methods

Public Function Create(ByVal currwb As Workbook, _
                       ByVal trads As ITranslation, ByVal fObj As Object)

  checkValidity currwb, trads, fObj



End Function





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


Private Sub CreateExport(ByVal scope As Byte)

    BusyApp cursor:=xlWait
    
    'Add Error management
    On Error GoTo ErrHand

    InitializeTrads
    AskFilter tradmess
    Set expOut = OutputSpecs.Create(currwb, scope)
    expOut.Save tradmess, useFilter
    NotBusyApp
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox tradmess.TranslatedValue("MSG_ErrHandExport"), _
            vbOKOnly + vbCritical, _
            tradmess.TranslatedValue("MSG_Error")
    'Close all oppened workbooks
    expOut.CloseAll
    On Error GoTo 0
    NotBusyApp
End Sub