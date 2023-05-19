Attribute VB_Name = "FormLogicExport"
Attribute VB_Description = "Form implementation of Exports"


Option Explicit

Private Sub CMD_Export2_Click()

    Call export(2)

End Sub

Private Sub CMD_Export1_Click()

    Call export(1)

End Sub

Private Sub CMD_Export3_Click()

    Call export(3)

End Sub

Private Sub CMD_Export4_Click()

    Call export(4)

End Sub

Private Sub CMD_Export5_Click()

    Call export(5)

End Sub

Private Sub CMD_NouvCle_Click()

    Call NewKey

End Sub

Private Sub CMD_Retour_Click()

    F_Export.Hide

End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 200
    Me.height = 400

End Sub
