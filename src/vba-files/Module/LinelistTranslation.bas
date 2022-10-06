Attribute VB_Name = "LinelistTranslation"
Option Explicit
Option Private Module

Function GetLanguageCode(sString As String) As String
    Dim T_data As BetterArray                    'array of languages
    Dim T_codes As BetterArray                   'array of languages codes
    Dim T_values As BetterArray                  'values of languages

    Set T_data = New BetterArray
    T_data.LowerBound = 1
    Set T_codes = New BetterArray
    T_codes.LowerBound = 1
    Set T_values = New BetterArray
    T_values.LowerBound = 1

    GetLanguageCode = ""

    T_data.FromExcelRange SheetLLTranslation.ListObjects(C_sTabLLLang).DataBodyRange 'Language table
    T_values.Items = T_data.ExtractSegment(ColumnIndex:=1) 'language values
    T_codes.Items = T_data.ExtractSegment(ColumnIndex:=2) 'language codes

    If T_values.Includes(sString) Then
        'The Language code
        GetLanguageCode = T_codes.Item(T_values.IndexOf(sString))
    End If

End Function

Sub TranslateForm(UserFrm As UserForm)
    'management of the translation of the form captions

    Dim i As Integer
    Dim cControl As Control

    For Each cControl In UserFrm.Controls
        If TypeOf cControl Is MSForms.CommandButton Or (TypeOf cControl Is MSForms.Label) Or (TypeOf cControl Is MSForms.OptionButton) _
        Or (TypeOf cControl Is MSForms.Page) Or (TypeOf cControl Is MSForms.MultiPage) Or (TypeOf cControl Is MSForms.Frame) Or (TypeOf cControl Is MSForms.CheckBox) Then
            If TypeOf cControl Is MSForms.MultiPage Then
                For i = 0 To cControl.Pages.Count - 1
                    If cControl.Name = "MultiPage1" Then UserFrm.MultiPage1.Pages(i).Caption = LineListTranslatedValue(UserFrm.MultiPage1.Pages(i).Name, C_sTabTradLLForms)
                    If cControl.Name = "MultiPage2" Then UserFrm.MultiPage2.Pages(i).Caption = LineListTranslatedValue(UserFrm.MultiPage2.Pages(i).Name, C_sTabTradLLForms)
                Next i
            Else
                If Trim(cControl.Caption) <> "" Then cControl.Caption = LineListTranslatedValue(cControl.Name, C_sTabTradLLForms)
            End If
        End If
    Next cControl
End Sub

'Find correponding values in one listobject of the linelist translation sheet and translate them

Function LineListTranslatedValue(sText As String, sRngName As String)
    'Management of the translation of the Linelist

    Dim sLanguage As String
    Dim iNumCol As Integer
    Dim HeadersData As BetterArray
    Dim TransWksh As Worksheet
    Dim rng As Range

    Set HeadersData = New BetterArray
    Set TransWksh = ThisWorkbook.Worksheets(C_sSheetLLTranslation)
    Set rng = TransWksh.ListObjects(sRngName).Range

    LineListTranslatedValue = vbNullString

    HeadersData.FromExcelRange TransWksh.ListObjects(sRngName).HeaderRowRange
    sLanguage = TransWksh.Range(C_sRngLLLanguageCode)
    iNumCol = HeadersData.IndexOf(sLanguage)

    On Error Resume Next

    If iNumCol > 0 Then
        LineListTranslatedValue = Application.WorksheetFunction.VLookup(sText, rng, iNumCol, False)
    End If

    On Error GoTo 0
End Function

'Translate a message in the linelist (corresponding to the choosen language)
Function TranslateLLMsg(sMsgCode As String) As String
    TranslateLLMsg = LineListTranslatedValue(sMsgCode, C_sTabTradLLMsg)
End Function


