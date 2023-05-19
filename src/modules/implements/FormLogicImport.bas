Attribute VB_Name = "FormLogicImport"
Attribute VB_Description = "Form implementation of Import"
Option Explicit


Private Sub CMD_ClearData_Click()
    Call ControlClearData
End Sub

Private Sub CMD_ClearGeo_Click()
    Call ClearHistoricGeobase
End Sub

Private Sub CMD_ImportData_Click()
    Call ImportMigrationData
End Sub

Private Sub CMD_ImportGeo_Click()
    Call ImportGeobase
End Sub

Private Sub CMD_ImportGeoHistoric_Click()
    Call ImportHistoricGeobase
End Sub

Private Sub CMD_ImportMigQuit_Click()
    Me.Hide
End Sub

Private Sub CMD_ImportMigRep_Click()
    Me.Hide
    Call ShowImportReport
End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 200
    Me.height = 350

End Sub
