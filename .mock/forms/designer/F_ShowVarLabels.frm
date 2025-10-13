VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowVarLabels 
   Caption         =   "UserForm1"
   ClientHeight    =   7170
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13935
   OleObjectBlob   =   "F_ShowVarLabels.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ShowVarLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
    Me.height = 600
End Sub
