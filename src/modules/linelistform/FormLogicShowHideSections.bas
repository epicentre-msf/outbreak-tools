Attribute VB_Name = "FormLogicShowHideSections"
Attribute VB_Description = "Form code-behind for F_ShowHideSections"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHideSections")
'@depends LinelistEventsManager, EventLinelist, LLTranslation, TranslationObject

' This module is the complete code-behind of the F_ShowHideSections form and is
' copied into the form at deployment. The form hides and shows whole sections of
' the data entry sheet the show/hide form is open on, which is the same thing
' the section button does for the one section the cursor stands on.
'
' The controls the form carries:
'
'   LST_Sections   two columns, the section title and its status
'   OPT_Show       show the picked section
'   OPT_Hide       hide the picked section
'   CMD_Back       leave the form
'
' The two option buttons carry the names the show/hide form uses for the same
' two words, so the translation rows they read already exist.

Option Explicit


Private tradform As TranslationObject


'The translation helper is the one EventLinelist holds.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    If Not tradform Is Nothing Then Exit Sub

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicShowHideSections", _
                  "This linelist carries no usable translation sheet"

    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub LST_Sections_Click()
    ClickListShowHideSections Me.LST_Sections.ListIndex
End Sub

Private Sub OPT_Show_Click()
    ClickOptionsShowHideSections Me.LST_Sections.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHideSections Me.LST_Sections.ListIndex
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub

'Go back to show/hide. The form hides itself and asks for the parent back;
'ClickShowHide shows its form again once this form's opener has returned and
'saved. Calling ClickShowHide from here ran it inside that opener, before its
'save, and the sheet was put back to the state from before this form opened.
Private Sub LBL_Previous_Click()
    Me.Hide
    ClickShowHidePrevious
End Sub