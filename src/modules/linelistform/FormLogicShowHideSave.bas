Attribute VB_Name = "FormLogicShowHideSave"
Attribute VB_Description = "Form code-behind for F_ShowHideSave"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHideSave")

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"

Private tradform As TranslationObject
Private tradmess As TranslationObject


Private Sub InitializeTrads()
    Dim lltrads As LLTranslation

    If Not tradmess Is Nothing Then Exit Sub

    Set lltrads = LLTranslation.Create(ThisWorkbook.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
End Sub

'The saved choices of this workbook. Answers Nothing when the workbook carries
'no show/hide worksheet, which is what an older linelist looks like.
Private Function StoreOf() As ShowHideStore
    On Error Resume Next
    Set StoreOf = ShowHideStore.Create(ThisWorkbook)
    On Error GoTo 0
End Function

'The protection keys of the running linelist
Private Function PasswordsOf() As Passwords
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set PasswordsOf = linelistEvents.PasswordManager()
End Function

'Put the saved layout names in the list, in the order they were first saved
Private Sub RefreshLayoutList()
    Dim store As ShowHideStore
    Dim layoutNames As BetterArray
    Dim counter As Long

    Me.LST_LayoutSave.Clear

    Set store = StoreOf()
    If store Is Nothing Then Exit Sub

    Set layoutNames = store.LayoutNames()
    For counter = layoutNames.LowerBound To layoutNames.UpperBound
        Me.LST_LayoutSave.AddItem layoutNames.Item(counter)
    Next
End Sub

'The layout name the user has picked in the list
Private Function SelectedLayoutName() As String
    If Me.LST_LayoutSave.ListIndex < 0 Then Exit Function
    SelectedLayoutName = CStr(Me.LST_LayoutSave.List(Me.LST_LayoutSave.ListIndex))
End Function

'Tell the user to pick a layout in the list first
Private Sub WarnPickALayout()
    MsgBox tradmess.TranslatedValue("MSG_LayoutSelect"), _
           vbOKOnly + vbInformation, _
           tradmess.TranslatedValue("MSGB_Warning")
End Sub

'Save what every layered worksheet shows under a name the user types in
Private Sub CMD_ShowHideAdd_Click()
    Dim store As ShowHideStore
    Dim layoutName As String

    InitializeTrads

    layoutName = Trim$(InputBox(tradmess.TranslatedValue("MSG_LayoutName"), _
                                tradmess.TranslatedValue("MSG_Enter")))
    If LenB(layoutName) = 0 Then Exit Sub

    Set store = StoreOf()
    If store Is Nothing Then Exit Sub

    'A name the store already holds is overwritten once the user agrees. A new
    'name past the cap is refused here, where the user can be told why, rather
    'than silently inside the walk.
    If store.HasLayout(layoutName) Then
        If MsgBox(tradmess.TranslatedValue("MSG_LayoutReplace"), _
                  vbYesNo + vbExclamation, _
                  tradmess.TranslatedValue("MSGB_Warning")) = vbNo Then Exit Sub
    ElseIf store.LayoutCount() >= store.MaxSavedLayouts Then
        MsgBox tradmess.TranslatedValue("MSG_LayoutCap"), _
               vbOKOnly + vbExclamation, _
               tradmess.TranslatedValue("MSGB_Warning")
        Exit Sub
    End If

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState
    HandleSaveShowHideLayout ThisWorkbook, PasswordsOf(), layoutName

ErrHand:
    LinelistEventsManager.LLExitBusyState
    RefreshLayoutList
End Sub

'Put every layered worksheet in the state the picked layout describes
Private Sub CMD_ShowHideApply_Click()
    Dim layoutName As String

    InitializeTrads

    layoutName = SelectedLayoutName()
    If LenB(layoutName) = 0 Then
        WarnPickALayout
        Exit Sub
    End If

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState
    HandleRestoreShowHideLayout ThisWorkbook, PasswordsOf(), layoutName

ErrHand:
    LinelistEventsManager.LLExitBusyState
End Sub

'Drop the picked layout, once the user confirms
Private Sub CMD_ShowHideDelete_Click()
    Dim store As ShowHideStore
    Dim layoutName As String

    InitializeTrads

    layoutName = SelectedLayoutName()
    If LenB(layoutName) = 0 Then
        WarnPickALayout
        Exit Sub
    End If

    If MsgBox(tradmess.TranslatedValue("MSG_LayoutDelete") & " (" & layoutName & ")", _
              vbYesNo + vbExclamation, _
              tradmess.TranslatedValue("MSGB_Warning")) = vbNo Then Exit Sub

    Set store = StoreOf()
    If store Is Nothing Then Exit Sub

    store.DeleteLayout layoutName
    RefreshLayoutList
End Sub

'Rename the layout the user double-clicks, keeping its rows
Private Sub LST_LayoutSave_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim store As ShowHideStore
    Dim oldName As String
    Dim newName As String

    InitializeTrads

    oldName = SelectedLayoutName()
    If LenB(oldName) = 0 Then Exit Sub

    newName = Trim$(InputBox(tradmess.TranslatedValue("MSG_LayoutNewName"), _
                             tradmess.TranslatedValue("MSG_Enter"), oldName))
    If LenB(newName) = 0 Then Exit Sub
    If StrComp(newName, oldName, vbBinaryCompare) = 0 Then Exit Sub

    Set store = StoreOf()
    If store Is Nothing Then Exit Sub

    'Renaming onto another saved name would raise inside the store. Recasing
    'the same name is the one collision the store accepts, so it passes.
    If store.HasLayout(newName) And _
       (StrComp(newName, oldName, vbTextCompare) <> 0) Then
        MsgBox tradmess.TranslatedValue("MSG_LayoutExists"), _
               vbOKOnly + vbExclamation, _
               tradmess.TranslatedValue("MSGB_Warning")
        Exit Sub
    End If

    store.RenameLayout oldName, newName
    RefreshLayoutList
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub

'The store can change between two shows of the form, so the list is read
'again each time rather than once at initialize
Private Sub UserForm_Activate()
    RefreshLayoutList
End Sub
