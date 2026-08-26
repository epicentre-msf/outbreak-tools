Attribute VB_Name = "FormLogicShowHide"
Attribute VB_Description = "Form code-behind for F_ShowHideLL"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHideLL")
'@depends LinelistEventsManager, EventLinelist, LLTranslation, TranslationObject

' This module is the complete code-behind of the F_ShowHideLL form and is copied
' into the form at deployment. The form shows the variables of one data entry
' sheet and lets the user hide the ones the sheet does not need.
'
' The controls the form carries:
'
'   LST_LLVarNames         three columns, the label, the variable and its status
'   OPT_Show               show the picked variable
'   OPT_Hide               hide the picked variable
'   LBL_ID                 header of the variable column
'   LBL_Status             header of the status column
'   CMD_Back               leave the form
'   CMD_ShowHideLayout     open the saved layouts form
'   CMD_ShowHideMinimal    hide every optional variable at once
'   CMD_ShowHideSections   open the sections form
'
' Every one of them has a row in T_TradLLForms under its own name, so
' UserForm_Initialize translates the whole form in one call.

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
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicShowHide", _
                  "This linelist carries no usable translation sheet"

    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub

Private Sub LST_LLVarNames_Click()
    ClickListShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Show_Click()
    ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub CMD_ShowHideLayout_Click()
    Me.Hide
    ClickShowHideLayouts
End Sub

Private Sub CMD_ShowHideMinimal_Click()
    ClickShowHideMinimal
    Me.Hide
End Sub

'Open the sections form on top of this one. This form stays up, and the list
'behind is filled again once the sections form goes down, because a whole
'section may have moved while it was open.
Private Sub CMD_ShowHideSections_Click()
    Me.Hide
    ClickOpenShowHideSections
End Sub
