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
    Dim geo As ILLGeo
    

    'If does not have a name property do nothing
    '    Dim sIsoCountry As String, sCountry As String, sSubCounty As String, sWard As String, sPlace As String, sFacility As String

    On Error GoTo no
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.EnableEvents = False

    If Target.Name.Name = "RNG_LangDesigner" Then
        TranslateDesignerMain
    ElseIf Target.Name.Name = "RNG_LLForm" Then
        
        'Language of LinelistForms
        SheetLLTranslation.Range("RNG_LLLanguage").Value = Target.Value
        SheetLLTranslation.Calculate

        'Language Code in the Geo Sheet
        Set geo = LLGeo.Create(SheetGeo)
        geo.Translate rawNames:=True
        SheetGeo.Range("RNG_GeoLangCode").Value = SheetLLTranslation.Range("RNG_LLLanguageCode").Value
        SheetGeo.Calculate

    ElseIf Target.Name.Name = "RNG_LLPassword" Then
    
        SheetPassword.Range("RNG_DebuggingPassword").Value = Target.Value

    ElseIf Target.Name.Name = "RNG_LangSetup" Then
    
        SheetLLTranslation.Range("RNG_DictionaryLanguage").Value = Target.Value
        SheetGeo.Range("RNG_MetaLang").Value = Target.Value
    End If

    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.EnableEvents = True
    Exit Sub
no:
    Application.EnableEvents = True
End Sub

