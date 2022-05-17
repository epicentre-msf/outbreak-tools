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

    sLanguage = Application.WorksheetFunction.VLookup(Sheets("linelist-translation").[RNG_LLLanguage].value, Sheets("linelist-translation").[T_Lang2], 2, False)

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

Sub ImportLanguage(sPath As String)
'Import languages from the setup file and sheet Translation

    Dim sAdr1 As String, sAdr2 As String
    Dim Wkb As Workbook
    Dim src As Range 'Source range
    Dim dest As Range 'Destination range

    
    With SheetDesTranslation
        .Range(.Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").Column), _
               .Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").End(xlToRight).Column)).ClearContents
    End With

    SheetSetTranslation.Cells.Clear
    
    BeginWork xlsapp:=Application
    
    Set Wkb = Workbooks.Open(Filename:=sPath)
    Set src = Wkb.Worksheets("Translations").ListObjects("Tab_Translations").HeaderRowRange
    Set dest = SheetDesTranslation.Range("T_Lst_Lang")

    src.Copy dest
    
    Set src = Wkb.Worksheets("Translations").ListObjects("Tab_Translations").Range
    Set dest = SheetSetTranslation.Range("A" & C_eStartlinestransdata)
    
    src.Copy dest

    sAdr1 = SheetDesTranslation.Range("T_Lst_Lang").Address
    sAdr2 = SheetDesTranslation.Range("T_Lst_Lang").End(xlToRight).Address

    SheetMain.Range(C_sRngLangSetup).value = ""
    
    
    With SheetMain.Range(C_sRngLangSetup).Validation
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
    
    Wkb.Close
    Set Wkb = Nothing

    EndWork xlsapp:=Application
End Sub

Function F_TransLL_Create(ByVal sText As String, iNumCol As Integer, Optional sType As String)
'translation of the linelist according to the chosen dictionary language

    Dim i As Integer, iStart As Integer
    Dim sFormula As String, sLabelTranlate As String

    If SheetMain.[RNG_LangSetup].value = "" Then
        F_TransLL_Create = sText
        Exit Function
    End If

    If iNumCol = 2 Or sText = "" Then
        F_TransLL_Create = sText
    Else
        If sType = "Formula" Then
            sFormula = sText
            sFormula = Replace(sFormula, Chr(34) & Chr(34), "")
            If InStr(1, sFormula, Chr(34), 1) > 0 Then
                For i = 1 To Len(sFormula)
                    If Mid(sFormula, i, 1) = Chr(34) Then
                        If iStart = 0 Then
                            iStart = i + 1
                        Else
                            sLabelTranlate = Application.WorksheetFunction.VLookup(Mid(sFormula, iStart, i - iStart), Sheets("Translation").[Tab_Translations].value, iNumCol - 1, False)
                            If sLabelTranlate <> "" Then sText = Replace(sText, Mid(sFormula, iStart, i - iStart), sLabelTranlate)
                            iStart = 0
                        End If
                    End If
                Next i
                F_TransLL_Create = sText
            Else
                F_TransLL_Create = sText
            End If
        Else
            If Application.WorksheetFunction.VLookup(sText, Sheets("Translation").[Tab_Translations].value, iNumCol - 1, False) <> "" Then
                F_TransLL_Create = Application.WorksheetFunction.VLookup(sText, Sheets("Translation").[Tab_Translations].value, iNumCol - 1, False)
            Else
                F_TransLL_Create = sText
            End If
        End If
    End If

End Function
