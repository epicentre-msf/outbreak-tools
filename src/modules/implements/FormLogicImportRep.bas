Attribute VB_Name = "FormLogicImportRep"
Attribute VB_Description = "Form implementation of Reports on import"
Option Explicit

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 550
    Me.height = 450

End Sub

Private Sub CMD_ImpRepQuit_Click()
    Me.Hide
End Sub