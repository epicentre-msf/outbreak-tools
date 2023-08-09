Attribute VB_Name = "FormLogicShowHide"
Attribute VB_Description = "Events show/hide in the linelist"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Events show/hide in the linelist")

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
