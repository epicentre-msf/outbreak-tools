Attribute VB_Name = "FormLogicExport"
Attribute VB_Description = "Form implementation of Exports"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form implementation of Exports")

Option Explicit

Private Const PASSWORDSHEET As String = "__pass"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private pass As ILLPasswords
Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation 'Translation of messasges


'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim passsh As Worksheet
    Dim currwb As Workbook

    Set currwb = ThisWorkbook
    Set lltranssh = currwb.Worksheets(LLSHEET)
    Set dicttranssh = currwb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set passsh = currwb.Worksheets(PASSWORDSHEET)
    Set pass = LLPasswords.Create(passsh)
End Sub


Private Sub CMD_NewKey_Click()
    InitializeTrads
    pass.GenerateKey tradmess
End Sub

Private Sub CMD_ShowKey_Click()
    InitializeTrads
    pass.DisplayPrivateKey tradmess
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

'Translate the form, add form sizes.
Private Sub UserForm_Initialize()
    'Manage language
    
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 200
    Me.height = 400
End Sub
