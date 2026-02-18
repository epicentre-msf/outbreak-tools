Attribute VB_Name = "FormLogicImportRep"
Attribute VB_Description = "Form code-behind for F_ImportRep"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ImportRep")

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"

Private tradform As ITranslationObject


Private Sub InitializeTrads()
    Dim lltrads As ILLTranslation

    Set lltrads = LLTranslation.Create(ThisWorkbook.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 550
    Me.Height = 450
End Sub

Private Sub CMD_ImpRepQuit_Click()
    Me.Hide
End Sub

Private Sub LBL_Previous_Click()
    Me.Hide
    F_Advanced.Show
End Sub
