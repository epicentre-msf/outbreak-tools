VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ImportRep 
   Caption         =   "Import Summary"
   ClientHeight    =   5370
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10320
   OleObjectBlob   =   "F_ImportRep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ImportRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form implementation of Reports on import")

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

'Set form width and height, add translations
Private Sub UserForm_Initialize()

    'Manage language
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 550
    Me.height = 450
End Sub

Private Sub CMD_ImpRepQuit_Click()
    Me.Hide
End Sub

Private Sub LBL_Previous_Click()
    Me.Hide
    F_Advanced.Show
End Sub
