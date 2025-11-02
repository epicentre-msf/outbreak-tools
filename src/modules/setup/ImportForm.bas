Attribute VB_Name = "ImportForm"
Attribute VB_Description = "Imports form logics"

'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString
Option Explicit
Private NumberOfClicks As Long
Private Const LimitOfClicks As Long = 10


Private Sub DictionaryCheck_Click()
    ChoiceCheck.Value = DictionaryCheck.Value
    ExportsCheck.Value = DictionaryCheck.Value
End Sub

Private Sub DoButton_Click()
    'Check if everything is fine with the setup and import one
    SetupHelpers.ImportOrCleanSetup
End Sub

Private Sub LoadButton_Click()
    'Load a new setup
   Dim filePath As String

   filePath = SetupHelpers.SelectSetupImportPath("*.xlsb")
   If LenB(filePath) <> 0 Then [Imports].LabPath.Caption = "Path: " & filePath
End Sub

Private Sub Quit_Click()
    Me.LabProgress.Caption = vbNullString
    Me.Hide
End Sub

Private Sub UserForm_Click()
    NumberOfClicks = NumberOfClicks + 1

    If NumberOfClicks = 9 Then
        Me.LabProgress.Caption = "click somewhere in the form again to enter debug mode"
    End If

    Dim pass As IPasswords
    Dim pwdUser As String

    'Write prompt for the debugging password


    Set pass = SetupHelpers.ResolveSetupPassword()
    If pwdUser = pass.Value("debuggingpassword") Then
        pass.EnterDebugMode
        Me.LabProgress.Caption = vbNullString
        MsgBox "Setup in debug mode!"
        Me.Hide
    End If

End Sub

