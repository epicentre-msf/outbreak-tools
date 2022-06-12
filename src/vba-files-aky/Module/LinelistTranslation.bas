Attribute VB_Name = "LinelistTranslation"
Option Explicit

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

    Dim sLanguage As String
    Dim i As Integer
    Dim cControl As Control

    For Each cControl In UserFrm.Controls
        If TypeOf cControl Is MSForms.CommandButton Or (TypeOf cControl Is MSForms.Label) Or (TypeOf cControl Is MSForms.OptionButton) _
        Or (TypeOf cControl Is MSForms.Page) Or (TypeOf cControl Is MSForms.MultiPage) Or (TypeOf cControl Is MSForms.Frame) Or (TypeOf cControl Is MSForms.CheckBox) Then
            If TypeOf cControl Is MSForms.MultiPage Then
                For i = 0 To cControl.Pages.Count - 1
                    If cControl.Name = "MultiPage1" Then UserFrm.MultiPage1.Pages(i).Caption = TranslateLineList(UserFrm.MultiPage1.Pages(i).Name, C_sTabTradLLForms)
                    If cControl.Name = "MultiPage2" Then UserFrm.MultiPage2.Pages(i).Caption = TranslateLineList(UserFrm.MultiPage2.Pages(i).Name, C_sTabTradLLForms)
                Next i
            Else
                If Trim(cControl.Caption) <> "" Then cControl.Caption = TranslateLineList(cControl.Name, C_sTabTradLLForms)
            End If
        End If
    Next cControl
End Sub


Function TranslateLineList(sText As String, sRngName As String)
    'Management of the translation of the Linelist

    Dim sLanguage As String
    Dim iNumCol As Integer
    Dim HeadersData As BetterArray
    Dim TransWksh As Worksheet
    Dim Rng As Range

    Set HeadersData = New BetterArray
    Set TransWksh = ThisWorkbook.Worksheets(C_sSheetLLTranslation)
    Set Rng = TransWksh.ListObjects(sRngName).Range

    TranslateLineList = vbNullString

    HeadersData.FromExcelRange TransWksh.ListObjects(sRngName).HeaderRowRange
    sLanguage = TransWksh.Range(C_sRngLLLanguageCode)
    iNumCol = HeadersData.IndexOf(sLanguage)

    On Error Resume Next

    If iNumCol > 0 Then
         TranslateLineList = Application.WorksheetFunction.VLookup(sText, Rng, iNumCol, False)
    End If

    On Error GoTo 0
    Set HeadersData = Nothing
End Function

Sub ImportLangAnalysis(sPath As String)
'Import languages from the setup file and sheets Translation and Analysis

    Dim Wkb As Workbook
    Dim sAdr1 As String
    Dim sAdr2 As String
    Dim src As Range
    Dim dest As Range

    With SheetDesTranslation
        .Range(.Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").Column), _
               .Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").End(xlToRight).Column)).ClearContents
    End With

    SheetSetTranslation.Cells.Clear

    BeginWork xlsapp:=Application
    Application.EnableEvents = False
    Application.EnableAnimations = False

    Set Wkb = Workbooks.Open(Filename:=sPath)

    SheetSetTranslation.Cells.Clear

    'Copy the languages
    Set src = Wkb.Worksheets(C_sParamSheetTranslation).ListObjects(C_sTabTranslation).Range
    With SheetSetTranslation
        Set dest = .Range(.Cells(C_eStartLinesTransdata, 1), .Cells(C_eStartLinesTransdata + src.Rows.Count, src.Columns.Count))
        dest.value = src.value
        Set dest = .Range(.Cells(C_eStartLinesTransdata + 1, 1), .Cells(C_eStartLinesTransdata + src.Rows.Count, src.Columns.Count))
        .Listobjects.Add(xlSrcRange, dest, xlYes).Name = C_sTabTranslation
    End With

    'Now Add the list object

    sAdr1 = SheetDesTranslation.Range("T_Lst_Lang").Address
    sAdr2 = SheetDesTranslation.Range("T_Lst_Lang").End(xlToRight).Address

    Wkb.Close

    'Set Validation, 1 is Error
    Call Helpers.SetValidation(SheetMain.Range(C_sRngLangSetup), "='" & SheetDesTranslation.Name & "'!" & sAdr1 & ":" & sAdr2, 1)

    Set Wkb = Nothing
    Set src = Nothing
    Set dest = Nothing

    SheetMain.Range(C_sRngLangSetup).value = SheetSetTranslation.Cells(C_eStartLinesTransdata, 1).value
    EndWork xlsapp:=Application
    Application.EnableEvents = True
    Application.EnableAnimations = True
End Sub

'Translate a message in the linelist (corresponding to the choosen language)
Function TranslateLLMsg(sMsgCode As String) As String
    TranslateLLMsg = TranslateLineList(sMsgCode, C_sTabTradLLMsg)
End Function

'--------------- Writing functions to translate the dictionary and other parts -----------------------------------------

'A function to translate on column in one sheet

Function GetTranslatedValue(ByVal sText As String) As String

    GetTranslatedValue = vbNullString

    Dim iColLang As Integer
    Dim rngTrans As Range
    Dim sLangSetup As String
    Dim iRow As Integer

    'search in linelist language
    sLangSetup = SheetMain.Range(C_sRngLangSetup).value
    iColLang = IIf(sLangSetup <> "", SheetSetTranslation.Rows(C_eStartLinesTransdata).Find(What:=sLangSetup, LookAt:=xlWhole).Column, C_eStartcolumntransdata)

    With DesignerWorkbook.Worksheets(C_sParamSheetTranslation)
        Set rngTrans = .ListObjects(C_sTabTranslation).DataBodyRange
    End With

    On Error Resume Next
        iRow = rngTrans.Find(What:=sText, LookAt:=xlWhole).Row
        GetTranslatedValue = SheetSetTranslation.Cells(iRow, iColLang).value
    On Error GoTo 0

End Function

Sub TranslateColumn(iCol As Integer, sSheetName As String)
    Dim iLastRow As Integer
    Dim Wksh As Worksheet
    Dim i
    Dim sText As String
    Dim rngTrans As Range

    If iCol > 0 Then 'Be sure the column exists
        Set Wksh = DesignerWorkbook.Worksheets(sSheetName)
        iLastRow = Wksh.Cells(Rows.Count, 1).End(xlUp).Row

        i = 2

        Do While i <= iLastRow
            If Wksh.Cells(i, iCol).value <> vbNullString Then
                sText = Wksh.Cells(i, iCol).value
                sText = GetTranslatedValue(sText)
                If sText <> vbNullString Then
                    Wksh.Cells(i, iCol).value = sText
                End If
            End If
            i = i + 1
        Loop
    End If
End Sub

Function TranslateCellFormula(ByVal sFormText As String) As String
    Dim j As Integer
    Dim iStart As Integer
    Dim sText As String

    Dim sFormula As String
    Dim sLabelTranlate As String

    TranslateCellFormula = vbNullString
    sText = sFormText

    iStart = 0

    sFormula = Replace(sText, Chr(34) & Chr(34), vbNullString)

    If InStr(1, sFormula, Chr(34), 1) > 0 Then
        For j = 1 To Len(sFormula)
            If Mid(sFormula, j, 1) = Chr(34) Then
                If iStart = 0 Then
                    iStart = j + 1
                Else
                    sLabelTranlate = GetTranslatedValue(Mid(sFormula, iStart, j - iStart))
                    If sLabelTranlate <> vbNullString Then
                        sText = Replace(sText, Mid(sFormula, iStart, j - iStart), sLabelTranlate)
                    End If
                    iStart = 0
                End If
            End If
        Next
    End If

    If sText <> vbNullString Then
        TranslateCellFormula = sText
    End If
End Function



Sub TranslateColumnFormula(iCol As Integer, sSheetName As String)

    Dim i As Integer
    Dim iLastRow As Integer
    Dim sText As String
    Dim Wksh As Worksheet

    Set Wksh = DesignerWorkbook.Worksheets(sSheetName)
    iLastRow = Wksh.Cells(Rows.Count, 1).End(xlUp).Row

    i = 2

    Do While i <= iLastRow
        sText = Wksh.Cells(i, iCol).value
        sText = TranslateCellFormula(sText)
        If sText <> vbNullString Then Wksh.Cells(i, iCol).value = sText
        i = i + 1
    Loop

End Sub


'A Function to translate the dictionary

'Translation of the dictionary

Sub TranslateDictionary()

    'List of columns to Translate
    Dim DictHeaders As BetterArray
    Dim iCol As Integer

    Set DictHeaders = New BetterArray
    Set DictHeaders = GetHeaders(DesignerWorkbook, C_sParamSheetDict, 1)

    'Translate different columns

    'Main label
    iCol = DictHeaders.IndexOf(C_sDictHeaderMainLab)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Sub-label
    iCol = DictHeaders.IndexOf(C_sDictHeaderSubLab)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Note
    iCol = DictHeaders.IndexOf(C_sDictHeaderNote)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Sheet Name
    iCol = DictHeaders.IndexOf(C_sDictHeaderSheetName)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Main Section
    iCol = DictHeaders.IndexOf(C_sDictHeaderMainSec)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Sub Section
    iCol = DictHeaders.IndexOf(C_sDictHeaderSubSec)
    Call TranslateColumn(iCol, C_sParamSheetDict)
    'Message
    iCol = DictHeaders.IndexOf(C_sDictHeaderMessage)
    Call TranslateColumn(iCol, C_sParamSheetDict)

    'Formula
    iCol = DictHeaders.IndexOf(C_sDictHeaderFormula)
    Call TranslateColumnFormula(iCol, C_sParamSheetDict)

    Set DictHeaders = Nothing


End Sub


'Translation of the choices
Sub TranslateChoices()

    Dim ChoiceHeaders As BetterArray
    Dim iCol As Integer

    Set ChoiceHeaders = New BetterArray
    Set ChoiceHeaders = GetHeaders(DesignerWorkbook, C_sParamSheetChoices, 1)

    'Label Short
    iCol = ChoiceHeaders.IndexOf(C_sChoiHeaderLabShort)
    Call TranslateColumn(iCol, C_sParamSheetChoices)

    'Label
    iCol = ChoiceHeaders.IndexOf(C_sChoiHeaderLab)
    Call TranslateColumn(iCol, C_sParamSheetChoices)

End Sub


'Translation of the Exports

Sub TranslateExports()

    'Second column is for label button (I hope)
    Call TranslateColumn(2, C_sParamSheetExport)

End Sub




Sub TranslateLinelistData()
'translation of the Export, Dictionary and Choice sheets for the linelist

    BeginWork xlsapp:=Application

    'Dictionary
    Call TranslateDictionary
    'Choices
    Call TranslateChoices
    'Exports
    Call TranslateExports

    'Analysis...


    EndWork xlsapp:=Application

End Sub

