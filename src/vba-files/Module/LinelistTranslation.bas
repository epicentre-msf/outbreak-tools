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

    'Copy the languages
    Set src = Wkb.Worksheets("Translations").ListObjects("Tab_Translations").HeaderRowRange
    Set dest = SheetDesTranslation.Range("T_Lst_Lang")
    src.Copy dest


    'Copy the translation data
    Set src = Wkb.Worksheets("Translations").ListObjects("Tab_Translations").Range
    Set dest = SheetSetTranslation.Range("A" & C_eStartlinestransdata)
    src.Copy dest

    sAdr1 = SheetDesTranslation.Range("T_Lst_Lang").Address
    sAdr2 = SheetDesTranslation.Range("T_Lst_Lang").End(xlToRight).Address
    
    'Set Validation, 1 is Error
    Call Helpers.SetValidation(SheetMain.Range(C_sRngLangSetup), "='" & SheetDesTranslation.Name & "'!" & sAdr1 & ":" & sAdr2, 1)

    Wkb.Close
    Set Wkb = Nothing
    Set src = Nothing
    Set dest = Nothing

    SheetMain.Range(C_sRngLangSetup).value = SheetSetTranslation.Cells(C_eStartlinestransdata, 1).value
    EndWork xlsapp:=Application
    Application.EnableEvents = True
    Application.EnableAnimations = True
End Sub

'Translate a message in the linelist (corresponding to the choosen language)
Function TranslateLLMsg(sMsgCode As String) As String
    TranslateLLMsg = TranslateLineList(sMsgCode, C_sTabTradLLMsg)
End Function

Sub Translate_Manage()
'translation of the Export, Dictionary and Choice sheets for the linelist

    Dim iCol As Integer, iStart As Integer, i As Integer, j As Integer, iColLang As Integer, iRow As Integer
    Dim iCptRow As Integer, iCptCol As Integer, iCptSheet As Integer
    Dim sText As String, sFormula As String, sLabelTranlate As String
    Dim arrColumn() As String
    Dim SheetActive As Worksheet

    Application.ScreenUpdating = False

    'search in linelist language
    iColLang = IIf([RNG_LangSetup].value <> "", SheetSetTranslation.Rows(4).Find(What:=SheetMain.[RNG_LangSetup].value, LookAt:=xlWhole).Column, 2)

'level sheet
    For iCptSheet = 1 To 3

        Select Case iCptSheet
            Case 1
                arrColumn = Split(sCstColDictionary, "|")
                Set SheetActive = DesignerWorkbook.Worksheets(C_sParamSheetDict)
            Case 2
                arrColumn = Split(sCstColChoices, "|")
                Set SheetActive = DesignerWorkbook.Worksheets(C_sParamSheetChoices)
            Case 3
                arrColumn = Split(sCstColExport, "|")
                Set SheetActive = DesignerWorkbook.Worksheets(C_sParamSheetExport)
        End Select

'***********************************************************************
'il faut virer les 2 lignes de codes suivantes *************************
'***********************************************************************

'        If SheetMain.[RNG_LangSetup].value = "" Then Exit For
'        If iColLang = 2 Then Exit For

        iCptRow = 1

        Do While SheetActive.Cells(iCptRow, 1).value <> ""
            iCptRow = iCptRow + 1
        Loop

    'level column
        For iCptCol = LBound(arrColumn, 1) To UBound(arrColumn, 1)

            If Not SheetActive.Rows(1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole) Is Nothing Then _
            iCol = SheetActive.Rows(1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole).Column

            i = 2
        'level Row
            Do While i < iCptRow
                If SheetActive.Cells(i, iCol).value <> "" Then
                    sText = SheetActive.Cells(i, iCol).value
                    If arrColumn(iCptCol) = "Formula" Then 'in case of formula
                        sFormula = sText
                        sFormula = Replace(sFormula, Chr(34) & Chr(34), "")
                        If InStr(1, sFormula, Chr(34), 1) > 0 Then
                            For j = 1 To Len(sFormula)
                                If Mid(sFormula, j, 1) = Chr(34) Then
                                    If iStart = 0 Then
                                        iStart = j + 1
                                    Else
                                        sLabelTranlate = Application.WorksheetFunction.VLookup(Mid(sFormula, iStart, j - iStart), SheetSetTranslation.[Tab_Translations].value, iColLang - 1, False)
                                        If sLabelTranlate <> "" Then sText = Replace(sText, Mid(sFormula, iStart, j - iStart), sLabelTranlate)
                                        iStart = 0
                                    End If
                                End If
                            Next j
                            SheetActive.Cells(i, iCol).value = sText
                         End If
                    Else
                        iRow = SheetSetTranslation.[Tab_Translations].Find(What:=sText, LookAt:=xlWhole).Row
                        If SheetSetTranslation.Cells(iRow, iColLang).value <> "" Then _
                        SheetActive.Cells(i, iCol).value = SheetSetTranslation.Cells(iRow, iColLang).value

                    End If
                End If
                i = i + 1
            Loop

        Next iCptCol

    Next iCptSheet

    Application.ScreenUpdating = True

End Sub

