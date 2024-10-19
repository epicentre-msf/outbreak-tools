Attribute VB_Name = "FormLogicEpiWeek"
Attribute VB_Description = "Events for epiweek start selection"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Events for epiweek start selection")

Option Explicit

Private Const UPDATESHEET As String = "updates__"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const RNGEPIWEEKSTART As String = "RNG_EpiWeekStart"



Private upobj As IUpVal
Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation
Private TriggerMode As Boolean


'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set lltranssh = wb.Worksheets(LLSHEET)
    Set dicttranssh = wb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set upobj = UpVal.Create(ThisWorkbook.Worksheets(UPDATESHEET))
End Sub

Private Sub RecomputeAndUpdate(ByVal startVal As Integer, ByVal captionValue As String)

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim tagValues As BetterArray
    Dim confirm As Integer
    
    'Ask for confirmation before proceeding
    confirm = MsgBox( _
        tradmess.TranslatedValue("MSG_ChangeStart") & Chr(10) & captionValue, _
        vbInformation + vbYesNo, _
        tradmess.TranslatedValue("MSG_Confirm") _
    )

    If confirm = vbNo Then GoTo Leave


    'Update the value on the update worksheet
    upobj.SetValue RNGEPIWEEKSTART, startVal

    Set wb = ThisWorkbook

    'Updating formulas in worksheets
    Set tagValues = New BetterArray
    tagValues.Push "HList", "VList", "TS-Analysis", "SP-Analysis", _
                   "Uni-Bi-Analysis", "SPT-Analysis"

    For Each sh in wb.Worksheets
        If tagValues.Includes(sh.Cells(1, 3).Value) Then

            On Error Resume Next
                sh.UsedRange.Calculate
            On Error GoTo 0
        End If
    Next

    MsgBox tradmess.TranslatedValue("MSG_Done")

Leave:
    Me.Hide
End Sub


Private Sub OptionMonday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 1, Me.OptionMonday.Caption
End Sub

Private Sub OptionTuesday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 2, Me.OptionTuesday.Caption
End Sub

Private Sub OptionWednesday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 3, Me.OptionWednesday.Caption
End Sub

Private Sub OptionThursday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 4, Me.OptionThursday.Caption
End Sub

Private Sub OptionFriday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 5, Me.OptionFriday.Caption
End Sub

Private Sub OptionSaturday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 6, Me.OptionSaturday.Caption
End Sub

Private Sub OptionSunday_Click()
    If Not TriggerMode Then Exit Sub
    RecomputeAndUpdate 0, Me.OptionSunday.Caption
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    'Manage language
    Me.Caption = tradform.TranslatedValue(Me.Name)

    tradform.TranslateForm Me

    Me.width = 170
    Me.height = 390
End Sub

'@EntryPoint
Public Sub ShowDefaultEpiWeek()

    InitializeTrads
    TriggerMode = False

    On Error GoTo ErrHand

    Select Case CLng(upobj.Value(RNGEPIWEEKSTART))
    Case 1
        Me.OptionMonday.Value = True
    Case 2
        Me.OptionTuesday.Value = True
    Case 3
        Me.OptionWednesday.Value = True
    Case 4
        Me.OptionThursday.Value = True
    Case 5
        Me.OptionFriday.Value = True
    Case 6
        Me.OptionSaturday.Value = True
    Case 0
        Me.OptionSunday.Value = True
    End Select

ErrHand:
    TriggerMode = True
    Me.Show
End Sub