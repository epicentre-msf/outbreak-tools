Attribute VB_Name = "ImportForm"
Attribute VB_Description = "Imports form logics"

'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString
Option Explicit
Private NumberOfClicks As Long
Private Const LimitOfClicks As Long = 5


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

Private Sub UserForm_Activate()
    NumberOfClicks = 0
End Sub

Private Sub LabPath_Click()
    NumberOfClicks = NumberOfClicks + 1

    If NumberOfClicks = (LimitOfClicks - 1) Then
        Me.LabProgress.Caption = "click somewhere in the box again to enter debug mode"
        Exit Sub
    End If

    If NumberOfClicks < LimitOfClicks Then Exit Sub

    Dim pass As IPasswords
    Dim pwdUser As Variant
    Dim expectedPassword As String
    Dim appState As IApplicationState

    Me.LabProgress.Caption = vbNullString

    Set pass = SetupHelpers.ResolveSetupPasswords()
    expectedPassword = pass.Value("debuggingpassword")

    pwdUser = Application.InputBox("Enter the debugging password.", _
                                   "Debugging Password", Type:=2)

    If (VarType(pwdUser) = vbBoolean) And (pwdUser = False) Then GoTo cleanExit

    If StrComp(CStr(pwdUser), expectedPassword, vbBinaryCompare) = 0 Then
        'Wrap EnterDebugMode in BusyState to suppress screen flickering
        'while iterating through all worksheets to unprotect them
        On Error GoTo DebugCleanup
        Set appState = ApplicationState.Create(Application)
        appState.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait
        pass.EnterDebugMode
        appState.Restore
        On Error GoTo 0

        Me.LabProgress.Caption = vbNullString
        MsgBox "Setup in debug mode!"
        Me.Hide
    Else
        Me.LabProgress.Caption = "Incorrect password."
    End If

cleanExit:
    NumberOfClicks = 0
    Exit Sub

DebugCleanup:
    On Error Resume Next
    If Not appState Is Nothing Then appState.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0
    Resume cleanExit
End Sub
