Attribute VB_Name = "FormLogicCustomFilters"
Attribute VB_Description = "Manage multiple filers in a linelist"

Option Explicit

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Manage multiple filers in a linelist")

Private Const PASSWORDSHEET As String = "__pass"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

Private pass As ILLPasswords
Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation 'Translation of messasges
Private custFiltObj As ICustomFilters
Private currwb As Workbook

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim passsh As Worksheet

    Set currwb = ThisWorkbook
    Set lltranssh = currwb.Worksheets(LLSHEET)
    Set dicttranssh = currwb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set passsh = currwb.Worksheets(PASSWORDSHEET)
    Set pass = LLPasswords.Create(passsh)

    Set custFiltObj = CustomFilters.Create()
End Sub

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub



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


'Translate the form, add form sizes.
Private Sub UserForm_Initialize()
    'Manage language
    
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 280
    Me.height = 450
End Sub
