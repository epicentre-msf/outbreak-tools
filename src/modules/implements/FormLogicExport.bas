Attribute VB_Name = "FormLogicExport"
Attribute VB_Description = "Form implementation of Exports"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable

Option Explicit

Private Const PASSWORDSHEET As String = "__pass"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private pass As ILLPasswords
Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation 'Translation of messasges
Private expOut As IOutputSpecs
Private currwb As Workbook
Private useFilter As Boolean

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim passsh As Worksheet

    Set currwb = ThisWorkbook
    Set lltranssh = currwb.Worksheets(LLSHEET)
    Set dicttranssh = currwb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set passsh = currwb.Worksheets(PASSWORDSHEET)
    Set pass = LLPasswords.Create(passsh)

End Sub

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Cursor = xlDefault
End Sub

Private Sub CreateExport(Byval scope As Byte)

    BusyApp cursor:=xlNorthwestArrow
    
    'Add Error management
    On Error GoTo errHand

    InitializeTrads
    AskFilter tradmess
    Set expOut = OutputSpecs.Create(currwb, scope)
    expOut.Save tradmess, useFilter
    NotBusyApp
    Exit Sub

errHand:
    On Error Resume Next
    MsgBox  tradmess.TranslatedValue("MSG_ErrHandExport"), _ 
            vbOKOnly + vbCritical, _ 
            tradmess.TranslatedValue("MSG_Error")
    'Close all oppened workbooks
    expOut.CloseAll
    On Error GoTo 0
    NotBusyApp
End Sub

Private Sub AskFilter(ByVal tradmess As ITranslation)
    
    Dim confirmFilterUse As Byte

    'Initialize the private useFilter
    useFilter = False

    If Me.CHK_ExportFiltered.Value Then

        confirmFilterUse = MsgBox(tradmess.TranslatedValue("MSG_AskFilter"), _ 
                                  vbYesNo + vbQuestion, _ 
                                  tradmess.TranslatedValue("MSG_ThereIsFilter"))

        If confirmFilterUse = vbYes Then
            'This function is in EventsGlobal Analysis, update filtertables will update all
            'filters in the current workbook.
            UpdateFilterTables calculate:=False
            useFilter = True
        Else
            Me.CHK_ExportFiltered.Value = False
        End If
    End If
End Sub


Private Sub CMD_Export1_Click()
    CreateExport 1
End Sub

Private Sub CMD_Export2_Click()
    CreateExport 2
End Sub

Private Sub CMD_Export3_Click()
    CreateExport 3
End Sub

Private Sub CMD_Export4_Click()
    CreateExport 4
End Sub

Private Sub CMD_Export5_Click()
    CreateExport 5
End Sub

Private Sub CMD_NewKey_Click()
    InitializeTrads
    pass.GenerateKey tradmess
End Sub

Private Sub CMD_ShowKey_Click()
    InitializeTrads
    pass.DisplayPrivateKey tradmess
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

'Translate the form, add form sizes.
Private Sub UserForm_Initialize()
    'Manage language
    
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 200
    Me.height = 400
End Sub
