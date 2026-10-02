Attribute VB_Name = "FormLogicShowHide"
Attribute VB_Description = "Form code-behind for F_ShowHideLL"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHideLL")

Option Explicit


Private Sub LST_LLVarNames_Click()
    ClickListShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Show_Click()
    ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub CMD_ShowHideMinimal_Click()
    ClickShowHideMinimal
    Me.Hide
End Sub
