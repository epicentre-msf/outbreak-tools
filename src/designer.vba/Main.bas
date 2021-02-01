Attribute VB_Name = "Main"
Option Explicit

Sub Compile()
    
    Dim oTranslations As Translations
    
    On Error GoTo Finally
    
    ProcessReset
    ProcessBegin
    
    Set oTranslations = New Translations
    
    oTranslations.FillFromRange Range("TABLE_TRANSLATE_START_CELL")
    
    oTranslations.AddTranslationFromRange Range("variables!B2:B34")
    oTranslations.AddTranslationFromRange Range("variables!C2:C34")
    oTranslations.AddTranslationFromRange Range("variables!G2:G34")
    oTranslations.AddTranslationFromRange Range("choices!D2:G34")
    
    oTranslations.FillRange Range("TABLE_TRANSLATE_START_CELL")
    
Finally:
    ProcessEnding
    
End Sub
