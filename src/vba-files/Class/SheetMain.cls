VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    'If does not have a name property do nothing
    '    Dim sIsoCountry As String, sCountry As String, sSubCounty As String, sWard As String, sPlace As String, sFacility As String

    On Error GoTo no
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.EnableEvents = False

    If Target.Name.Name = "RNG_LangDesigner" Then
        Call TranslateDesigner
    ElseIf Target.Name.Name = "RNG_LLForm" Then
        Call TranslateHeadGeo
    ElseIf Target.Name.Name = "RNG_LLPassword" Then
    
        SheetPassword.Range("RNG_DebuggingPassword").Value = Target.Value
    ElseIf Target.Name.Name = "RNG_LangSetup" Then
    
        SheetLLTranslation.Range("RNG_DictionaryLanguage").Value = Target.Value
    End If

    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.EnableEvents = True
    Exit Sub
no:
End Sub

