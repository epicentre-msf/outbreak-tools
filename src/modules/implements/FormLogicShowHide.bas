Attribute VB_Name = "FormLogicShowHide"
Attribute VB_Description = "Events show/hide in the linelist"
'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable

Option Explicit


Private Sub LST_LLVarNames_Click()
    ClickListShowHide (LST_LLVarNames.ListIndex)
End Sub

Private Sub OPT_Show_Click()
   ClickOptionsShowHide (LST_LLVarNames.ListIndex)
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide (LST_LLVarNames.ListIndex)
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub