Attribute VB_Name = "FormLogicShowVarLabels"
Attribute VB_Description = "Form code-behind for F_ShowVarLabels"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowVarLabels")

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"

Private tradform As ITranslationObject


Private Sub InitializeTrads()
    Dim lltrads As ILLTranslation

    Set lltrads = LLTranslation.Create(ThisWorkbook.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 700
    Me.Height = 380
End Sub
