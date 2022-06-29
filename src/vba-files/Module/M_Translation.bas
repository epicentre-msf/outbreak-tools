Attribute VB_Name = "M_Translation"
Option Explicit

Dim iRow As Integer
Dim iColStart As Integer
Dim iWrite As Integer

Sub Translate_Manage(Optional sType As String)
Attribute Translate_Manage.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim iCol As Integer
    Dim iLastRow As Integer
    Dim iRowStart As Integer
    Dim iStart As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iCptRow As Integer
    Dim iCptCol As Integer
    Dim iCptSheet As Integer
    Dim iCptTranslate() As Integer
    Dim sText As String, sMessage As String
    Dim arrColumn() As String
    Dim SheetActive As Worksheet

    Application.ScreenUpdating = False

    sheetTranslation.Unprotect C_sPassword

    iRow = sheetTranslation.[T_Translate_Start].Row
    iColStart = sheetTranslation.[T_Translate_Start].Column
    iLastRow = sheetTranslation.[Tab_Translations].Rows.Count + iRow
    iWrite = IIf(sheetTranslation.Cells(5, 2).Value = "", iLastRow, iLastRow + 1)

    'level sheet
    For iCptSheet = 1 To 4

        Select Case iCptSheet
            Case 1
                arrColumn = Split(sCstColDictionary, "|")
                Set SheetActive = sheetDictionary
                iRowStart = SheetActive.[T_Dictionary_Start].Row + 1
                iLastRow = SheetActive.[Tab_Dictionary].Rows.Count + 2
            Case 2
                arrColumn = Split(sCstColChoices, "|")
                Set SheetActive = SheetChoice
                iRowStart = SheetActive.[T_Choices_Start].Row + 1
                iLastRow = SheetActive.[Tab_Choices].Rows.Count + 1
            Case 3
                arrColumn = Split(sCstColExport, "|")
                Set SheetActive = sheetExport
                iRowStart = SheetActive.[T_Export_Start].Row + 1
                iLastRow = SheetActive.[Tab_Export].Rows.Count + 1
            Case 4
                arrColumn = Split(sCstColGlobalSummary, "|")
                Set SheetActive = sheetAnalysis
                iRowStart = 3
                iLastRow = 3 + SheetActive.ListObjects("Tab_global_summary").DataBodyRange.Rows.Count
        End Select

    'level column
        For iCptCol = LBound(arrColumn, 1) To UBound(arrColumn, 1)

            If Not SheetActive.Rows(iRowStart - 1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole) Is Nothing Then _
            iCol = SheetActive.Rows(iRowStart - 1).Find(What:=arrColumn(iCptCol), LookAt:=xlWhole).Column

            iCptRow = iRowStart

        'level Row
            Do While iCptRow <= iLastRow
                If SheetActive.Cells(iCptRow, iCol).Value <> "" Then
                    sText = SheetActive.Cells(iCptRow, iCol).Value
                    If arrColumn(iCptCol) = "Formula" Or arrColumn(iCptCol) = "Summary function" Then 'in case of formula
                        sText = Replace(sText, Chr(34) & Chr(34), "")
                        If InStr(1, sText, Chr(34), 1) > 0 Then
                            For i = 1 To Len(sText)
                                If Mid(sText, i, 1) = Chr(34) Then
                                    If iStart = 0 Then
                                        iStart = i + 1
                                    Else
                                        Call WriteTranslate(Mid(sText, iStart, i - iStart))
                                        iStart = 0
                                    End If
                                End If
                            Next i
                        End If
                    Else
                        Call WriteTranslate(sText)
                    End If
                End If
                iCptRow = iCptRow + 1
            Loop

        Next iCptCol

    Next iCptSheet

'removal of unnecessary labels
    iCptRow = sheetTranslation.[T_Translate_Start].Row + 1

    If sheetTranslation.[Tab_Translations].Columns.Count > 1 Then
        ReDim iCptTranslate(sheetTranslation.[Tab_Translations].Columns.Count - 2)
    Else
        ReDim iCptTranslate(sheetTranslation.[Tab_Translations].Columns.Count - 1)
    End If

    Do While iCptRow < iWrite
        If sheetTranslation.Cells(iCptRow, iColStart - 1).Value = "" Then
            sheetTranslation.Rows(iCptRow).Delete
            iWrite = iWrite - 1
        Else
            For i = 0 To UBound(iCptTranslate)
                If sheetTranslation.Cells(iCptRow, iColStart + (i + 1)).Value = "" Then iCptTranslate(i) = iCptTranslate(i) + 1
            Next i
            iCptRow = iCptRow + 1
        End If
    Loop

    iRow = sheetTranslation.[T_Translate_Start].Row + 1

    sheetTranslation.Select
    sheetTranslation.Range(Cells(iRow, iColStart - 1), Cells(iCptRow - 1, iColStart - 1)).ClearContents


'sheet Sorting
    sheetTranslation.Sort.SortFields.Clear
    sheetTranslation.Sort.SortFields.Add2 key:=sheetTranslation.[Tab_Translations_Sort], _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sheetTranslation.Sort
        .SetRange sheetTranslation.[Tab_Translations]
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Lock the first Column
Call LockFirstColumn
Call ProtectTranslationSheet

'elaboration of the user message

    j = 0

    For i = 0 To UBound(iCptTranslate)
        j = j + iCptTranslate(i)
    Next i

    If j = 0 Then
        sMessage = "The update of the translations sheet is correct."
        Application.ScreenUpdating = True
        bUpdate = False
        ActiveWorkbook.Save
        Exit Sub
    End If

    'if there are no columns to translate
    If sheetTranslation.Cells(iRow - 1, iColStart + 1).Value = "" Then
        ActiveWorkbook.Save
        Exit Sub
    End If

    For i = 0 To UBound(iCptTranslate)
        If iCptTranslate(i) > 0 Then _
        sMessage = sMessage & iCptTranslate(i) & " labels are missing for column " & _
        sheetTranslation.Cells(iRow - 1, iColStart + (i + 1)).Value & "." & Chr(10)
    Next i

    Application.ScreenUpdating = True

    If sType = "Close" Then
        Reponse = MsgBox(sMessage & Chr(10) & "Do you want to continue to close the workbook ?", vbYesNo, "verification of translation tags")
    Else
        MsgBox sMessage, vbCritical, "verification of translation tags"
        bUpdate = False
    End If

    'Lock the first column

    ActiveWorkbook.Save

End Sub

Sub WriteTranslate(sLabel As String)
'files the sheet translations

    If Not sheetTranslation.[Tab_Translations].Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        iRow = sheetTranslation.[Tab_Translations].Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=False).Row
        sheetTranslation.Cells(iRow, iColStart - 1).Value = 1
    Else
        sheetTranslation.Cells(iWrite, iColStart).Value = sLabel
        sheetTranslation.Cells(iWrite, iColStart - 1).Value = 1
        iWrite = iWrite + 1
    End If
End Sub

Sub LockFirstColumn()
    Dim iLastRow As Integer
    Dim rngFirstColumn As Range

    iLastRow = sheetTranslation.Cells(Rows.Count, 2).End(xlUp).Row

    With sheetTranslation
        Set rngFirstColumn = .Range(.Cells(5, 2), .Cells(iLastRow, 2))
    End With

    rngFirstColumn.Locked = True
End Sub


Sub ProtectTranslationSheet()
    sheetTranslation.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingColumns:=False, AllowFormattingRows:=False, _
        AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
End Sub
