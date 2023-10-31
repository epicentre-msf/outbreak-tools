Attribute VB_Name = "FormLogicShowVarLabels"
Attribute VB_Description = "Show the variables with corresponding labels in custom pivots tables"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Show the variables with corresponding labels in custom pivots tables")

Option Explicit

Private tradform As ITranslation   'Translation of forms
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet


    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

'Set form width and height, add translations
Private Sub UserForm_Initialize()

    'Manage language
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 650
    Me.height = 380
End Sub
