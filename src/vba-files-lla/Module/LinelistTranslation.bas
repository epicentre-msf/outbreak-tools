Attribute VB_Name = "LinelistTranslation"
Option Explicit

Sub TranslateForm(UserFrm As UserForm, rgPlage As Range)
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


Function translate_LineList(sText As String, rgPlage As Range)
'management of the translation of the Linelist Patient

    Dim sLanguage As String
    Dim iNumCol As Integer
    
    sLanguage = Application.WorksheetFunction.VLookup(Sheets("linelist-translation").[RNG_Language].value, _
    Sheets("linelist-translation").[T_Lang2], 2, False)

    Select Case sLanguage
        Case "ENG"
            iNumCol = 1
        Case "FRA"
            iNumCol = 2
        Case "POR"
            iNumCol = 3
        Case "ARA"
            iNumCol = 4
        Case "SPA"
            iNumCol = 5
    End Select
    
    translate_LineList = Application.WorksheetFunction.VLookup(sText, rgPlage, iNumCol, False)

End Function

Sub ImportLangAnalysis(sPath As String)
'Import languages from the setup file and sheets Translation and Analysis

    Dim sAdr1 As String, sAdr2 As String, sNomFic As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    SheetDesTranslation.Select
    SheetDesTranslation.[T_Lst_Lang].Select
    SheetDesTranslation.Range([T_Lst_Lang], Selection.End(xlToRight)).ClearContents
    
    SheetSetTranslation.Select
    SheetSetTranslation.Unprotect C_sDesignerPassword 'Default password for the designer
    Cells.Delete

    SheetAnalysis.Cells.Delete
    
    Workbooks.Open Filename:=sPath
    Sheets("Translations").Range("Tab_Translations[#Headers]").Copy
    
    sNomFic = Dir(sPath)
    
    DesignerWorkbook.Activate
    SheetDesTranslation.[T_Lst_Lang].PasteSpecial
    
    Windows(sNomFic).Activate
    Sheets("Translations").Select
    Cells.Copy
    
    DesignerWorkbook.Activate
    SheetSetTranslation.Range("A1").PasteSpecial
    SheetSetTranslation.Protect C_sDesignerPassword 'Default password for the designer
    
    Windows(sNomFic).Activate
    Sheets("Analysis").Select
    Cells.Copy

    DesignerWorkbook.Activate
    SheetAnalysis.Range("A1").PasteSpecial

    Windows(sNomFic).Close
       
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    sAdr1 = SheetDesTranslation.[T_Lst_Lang].Address
    sAdr2 = SheetDesTranslation.[T_Lst_Lang].End(xlToRight).Address
    
    SheetMain.Select
    SheetMain.[RNG_LangSetup].value = ""
    SheetMain.[RNG_LangSetup].Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & SheetDesTranslation.Name & "'!" & sAdr1 & ":" & sAdr2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    SheetMain.[RNG_LangSetup].value = SheetSetTranslation.Cells(4, 2).value
    
End Sub

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
                Sheets("Dictionary").Copy After:=Sheets(Sheets.Count)
                Set SheetActive = Sheets("Dictionary (2)")
                SheetActive.Name = "Dictionary_LL"
            Case 2
                arrColumn = Split(sCstColChoices, "|")
                Sheets("Choices").Copy After:=Sheets(Sheets.Count)
                Set SheetActive = Sheets("Choices (2)")
                SheetActive.Name = "Choices_LL"
            Case 3
                arrColumn = Split(sCstColExport, "|")
                Sheets("Exports").Copy After:=Sheets(Sheets.Count)
                Set SheetActive = Sheets("Exports (2)")
                SheetActive.Name = "Exports_LL"
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

