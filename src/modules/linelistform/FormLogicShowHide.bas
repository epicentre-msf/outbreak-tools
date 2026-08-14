Attribute VB_Name = "FormLogicShowHide"
Attribute VB_Description = "Form code-behind for F_ShowHideLL"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHideLL")

Option Explicit

'Register the state, and hide the form
Private Sub SaveAndHide()
    ClickShowHideMinimal
    Me.Hide
End Sub

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

Private Sub CMD_ShowHideLayout_Click()
    SaveAndHide
    ClickShowHideLayouts
End Sub

Private Sub CMD_ShowHideMinimal_Click()
    SaveAndHide
End Sub

'Open the sections form on top of this one. This form stays up, and the list
'behind is filled again once the sections form goes down, because a whole
'section may have moved while it was open.
Private Sub CMD_ShowHideSections_Click()
    SaveAndHide
    ClickOpenShowHideSections
End Sub
