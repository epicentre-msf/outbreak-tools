Attribute VB_Name = "LinelistTranslation"
Option Explicit

Sub TranslateForm(UserFrm As UserForm, rgPlage As Range) 'lla
'management of the translation of the form captions

    Dim sLanguage As String
    Dim iNumCol As Integer, i As Integer
    Dim cControl As Control
    
    sLanguage = Application.WorksheetFunction.VLookup(Sheets("linelist-translation").[RNG_Language].value, _
    Sheets("linelist-translation").[T_Lang2], 2, False)

    Select Case sLanguage
        Case "ENG"
            Exit Sub
        Case "FRA"
            iNumCol = 3
        Case "POR"
            iNumCol = 4
        Case "ARA"
            iNumCol = 5
        Case "SPA"
            iNumCol = 6
    End Select
    
    For Each cControl In UserFrm.Controls
        If TypeOf cControl Is MSForms.CommandButton Or (TypeOf cControl Is MSForms.Label) Or (TypeOf cControl Is MSForms.OptionButton) _
        Or (TypeOf cControl Is MSForms.Page) Or (TypeOf cControl Is MSForms.MultiPage) Or (TypeOf cControl Is MSForms.Frame) Then
            If TypeOf cControl Is MSForms.MultiPage Then
                For i = 0 To cControl.Pages.Count - 1
                    If cControl.Name = "MultiPage1" Then UserFrm.MultiPage1.Pages(i).Caption = _
                    Application.WorksheetFunction.VLookup(UserFrm.MultiPage1.Pages(i).Name, rgPlage, iNumCol, False)
                    If cControl.Name = "MultiPage2" Then UserFrm.MultiPage2.Pages(i).Caption = _
                    Application.WorksheetFunction.VLookup(UserFrm.MultiPage2.Pages(i).Name, rgPlage, iNumCol, False)
                Next i
            Else
                If Trim(cControl.Caption) <> "" Then _
                cControl.Caption = Application.WorksheetFunction.VLookup(cControl.Name, rgPlage, iNumCol, False)
            End If
        End If
    Next cControl
    
End Sub


Function translate_LineList(stext As String, rgPlage As Range) 'lla
'management of the translation of the Linelist Patient

    Dim sLanguage As String
    Dim iNumCol As Integer
    
    sLanguage = Application.WorksheetFunction.VLookup(Sheets("linelist-translation").[RNG_Language].value, _
    Sheets("linelist-translation").[T_Lang2], 2, False)

    Select Case sLanguage
        Case "ENG"
            Exit Function
        Case "FRA"
            iNumCol = 2
        Case "POR"
            iNumCol = 3
        Case "ARA"
            iNumCol = 4
        Case "SPA"
            iNumCol = 5
    End Select
    
    translate_LineList = Application.WorksheetFunction.VLookup(stext, rgPlage, iNumCol, False)

End Function



