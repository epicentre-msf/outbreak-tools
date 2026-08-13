Attribute VB_Name = "FormLogicShowVarLabels"
Attribute VB_Description = "Form code-behind for F_ShowVarLabels"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowVarLabels")

Option Explicit

Private tradform As TranslationObject


'The translation helper is the one EventLinelist holds. This module used to
'build its own on every open, and LLTranslation.Create validates all five
'translation tables per build.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicShowVarLabels", _
                  "This linelist carries no usable translation sheet"

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
