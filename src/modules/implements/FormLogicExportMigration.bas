Attribute VB_Name = "FormLogicExportMigration"
Attribute VB_Description = "Form implementation of Exports for Migration"
Option Explicit



Private Sub CMD_ExportMig_Click()
    Call ExportForMigration
End Sub

Private Sub CMD_ExportMigQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 200
    Me.height = 300

End Sub