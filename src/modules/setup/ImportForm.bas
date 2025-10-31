Attribute VB_Name = "ImportForm"
Attribute VB_Description = "Imports form logics"

'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString
Option Explicit


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
    [Imports].LabProgress.Caption = vbNullString
    Me.Hide
End Sub
