Attribute VB_Name = "FormLogicCustomFilters"
Attribute VB_Description = "Manage multiple filers in a linelist"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Manage multiple filers in a linelist")

Option Explicit

Private Sub LST_FiltersList_Click()
    Debug.Print Me.LST_FiltersList.ListIndex
End Sub

Private Sub CMD_ApplyFilter_Click()
   Debug.Print Me.LST_FiltersList.ListIndex
End Sub

Private Sub CMD_RemoveFilter_Click()
    Debug.Print Me.LST_FiltersList.ListIndex
End Sub

Private Sub CMD_RenameFilter_Click()

End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub