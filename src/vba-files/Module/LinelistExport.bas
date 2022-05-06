Attribute VB_Name = "LinelistExport"

Option Explicit

Private Function GetExportValues(iType As Byte, sSheetName As String) As BetterArray

    Dim ExportTableData As BetterArray   'Table of all the export data
    Dim ExportColumn As BetterArray 'one column of a data to export
    Dim SheetVarNamesData As BetterArray 'will contains all the variables for one sheet
    Dim VarNameData As BetterArray
    Dim YesNoExportData As BetterArray
    Dim ExportHeadersData As BetterArray
    Dim SheetNameData As BetterArray

    Dim sExportName As String 'the header of the export
    Dim i As Integer 'iterator

    Set SheetNameData = New BetterArray
    Set ExportTableData = New BetterArray
    Set ExportColumn = New BetterArray
    Set ExportHeadersData = New BetterArray
    Set SheetVarNamesData = New BetterArray
    Set VarNameData = New BetterArray
    Set YesNoExportData = New BetterArray

    ExportTableData.LowerBound = 1
    ExportColumn.LowerBound = 1
    SheetNameData.LowerBound = 1
    YesNoExportData.LowerBound = 1
    ExportHeadersData.LowerBound = 1
    SheetVarNamesData.LowerBound = 1
    VarNameData.LowerBound = 1

    'Get the type of the export
    Select Case iType
        Case 1
         sExportName = C_sDictHeaderExport1
        Case 2
         sExportName = C_sDictHeaderExport2
        Case 3
         sExportName = C_sDictHeaderExport3
        Case 4
         sExportName = C_sDictHeaderExport4
        Case 5
         sExportName = C_sDictHeaderExport5
        Case Else
         sExportName = C_sDictHeaderExport1
    End Select

    Set YesNoExportData = GetDictionaryColumn(sExportName)
    Set SheetNameData = GetDictionaryColumn(C_sDictHeaderSheetName)
    Set SheetVarNamesData = GetDictDataFromCondition(C_sDictHeaderSheetName, sSheetName, True)

    For i = 1 To YesNoExportData.UpperBound
        If YesNoExportData.Item(i) = "yes" And SheetNameData.Item(i) = sSheetName Then
            ExportHeadersData.Push ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(i + 1, 1).value 'Take the varname
        End If
    Next

    For i = 1 To ExportHeadersData.UpperBound
        If SheetVarNamesData.Includes(ExportHeadersData.Item(i)) Then
            With ThisWorkbook.Worksheets(sSheetName)
                    ExportColumn.FromExcelRange .Cells(C_eStartLinesLLData, SheetVarNamesData.IndexOf(ExportHeadersData.Item(i))), DetectLastColumn:=False, DetectLastRow:=True
                    ExportTableData.Item(i) = ExportColumn.Items
                    ExportColumn.Clear
            End With
        End If
    Next

    On Error GoTo errTranspose
    ExportTableData.ArrayType = BA_MULTIDIMENSION
    ExportTableData.Transpose
    Set GetExportValues = ExportTableData.Clone()
   
    Set ExportTableData = Nothing       'Table of all the export data
    Set ExportColumn = Nothing        'one column of a data to export
    Set SheetVarNamesData = Nothing   'will contains all the variables for one sheet
    Set VarNameData = Nothing
    Set YesNoExportData = Nothing
    Set ExportHeadersData = Nothing
    Set SheetNameData = Nothing

    Exit Function
errTranspose:
    MsgBox "Unable to transpose Export Table", vbOKOnly + vbCritical, "ERROR"
End Function


Sub Export(iTypeExport As Byte)

    Dim DictHeaders As BetterArray 'Headers of the dictionary
    Dim LLSheetData As BetterArray 'Vector of all sheets of type linelist
    Dim xlsapp As Excel.Application
    Dim DictData As BetterArray 'Values of the dictionary
    Dim ExportData As BetterArray
    Dim PathData As BetterArray 'Path to exports
    Dim VarNameData As BetterArray
    Dim ExportHeader As BetterArray
    Dim ChoicesData As BetterArray
    Dim TransData As BetterArray


    Dim i As Integer 'Iterator
    Dim sPrvSheetName As String
    Dim sPath As String
    Dim sDirectory As String
    Dim sPrevSheetName As String
    Dim sSheetName As String
    
    Set DictHeaders = New BetterArray
    Set DictData = New BetterArray
    Set LLSheetData = New BetterArray
    Set ExportData = New BetterArray
    Set PathData = New BetterArray
    Set VarNameData = New BetterArray
    
     DictHeaders.LowerBound = 1
     DictData.LowerBound = 1
     LLSheetData.LowerBound = 1
     ExportData.LowerBound = 1
     PathData.LowerBound = 1
     VarNameData.LowerBound = 1

    Set xlsapp = New Excel.Application

    'Get all the sheets of type Linelist
    Set DictHeaders = GetDictionaryHeaders()
    Set DictData = GetDictionaryData()
    Set VarNameData = GetDictionaryColumn(C_sDictHeaderVarName)
    Set ChoicesData = GetChoicesData()
    Set TransData = GetTransData()

    i = 1
    sPrevSheetName = ""
    While i <= DictData.Length
        If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetType))) = C_sDictSheetTypeLL Then
            If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))) <> sPrevSheetName Then
                sPrevSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                LLSheetData.Push sPrevSheetName
            End If
        End If
        i = i + 1
    Wend
    
    On Error GoTo exportErrHand
    With xlsapp
        .ScreenUpdating = False
        .Visible = False
        .Workbooks.Add

        'Adding the sheets for export
        .Worksheets(1).Name = C_sParamSheetDict

        'Writing the dictionary data
        DictHeaders.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
        DictData.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(2, 1)

        'Translation and choices sheets
        .Worksheets.Add(After:=.Worksheets(C_sParamSheetDict)).Name = C_sParamSheetChoices
        ChoicesData.ToExcelRange .Worksheets(C_sParamSheetChoices).Cells(1, 1)

        .Worksheets.Add(After:=.Worksheets(C_sParamSheetChoices)).Name = C_sParamSheetTranslation
        TransData.ToExcelRange .Worksheets(C_sParamSheetTranslation).Cells(1, 1)

        sPrevSheetName = C_sParamSheetDict

        i = 1
        While i <= LLSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = LLSheetData.Items(i)
            sPrevSheetName = LLSheetData.Items(i)
            ExportData.Clear

            Set ExportData = GetExportValues(iTypeExport, sPrevSheetName)
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(1, 1)
            i = i + 1
        Wend

        'pour l'enregistrement
        sPath = ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, 5).value

        If sPath <> "" Then
            PathData.Items = Split(sPath, "+")
            i = 1
            While i <= PathData.UpperBound
                PathData.Item(i) = Replace(xlsapp.WorksheetFunction.Trim(PathData.Items(i)), "+", "")

                If VarNameData.Includes(PathData.Items(i)) Then
                    sSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                    PathData.Item(i) = ThisWorkbook.Worksheets(sSheetName).Range(PathData.Items(i)).value
                Else
                    PathData.Item(i) = Replace(PathData.Items(i), Chr(34), "")
                End If
                i = i + 1
            Wend
           
            sPath = PathData.ToString(Separator:="_", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False) & _
                        "__" & Range("RNG_PublicKey").value & "__" & Format(Now, "yyyymmdd-HhNn")
            sDirectory = Helpers.LoadFolder
            If sDirectory <> "" Then
                sPath = sDirectory & Application.PathSeparator & sPath & ".xlsb"

                i = 0
                While Len(sPath) >= 255 And i < 3 'MSG_PathTooLong
                    MsgBox "The path of the export file is too long so the file name gets truncated. Please select a folder higher in the hierarchy to save the export (ex: Desktop, Downloads, Documents etc.)"
                    sDirectory = LoadFolder
                    If sDirectory <> "" Then
                        sPath = sDirectory & Application.PathSeparator & sPath
                    End If
                    i = i + 1
                Wend
                'on enregistre
                If i < 3 Then
                    .ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, CreateBackup:=False, Password:=Range("RNG_PrivateKey").value
                    MsgBox "File saved" & Chr(10) & "Password : " & Range("RNG_PrivateKey").value 'MSG_FileSaved        'MSG_Pass
                End If
                .ActiveWorkbook.Close
            End If
        End If
    End With

    xlsapp.Quit
    Set xlsapp = Nothing
    Set DictData = Nothing
    Set ChoicesData = Nothing
    Set TransData = Nothing

    Exit Sub
exportErrHand:
    MsgBox "Errors during export", vbOKOnly + vbCritical, "ERROR"
    Exit Sub
End Sub


Sub NewKey()
    '

    Dim nbLigne As Integer
    Dim T_Cle
    Dim i As Integer

    Sheets(C_sSheetPassword).Visible = xlSheetHidden

    T_Cle = Sheets(C_sSheetPassword).ListObjects(C_sTabkeys).DataBodyRange
    nbLigne = UBound(T_Cle, 1)

   ' Randomize
    i = Int(nbLigne * Rnd())
    Sheets(C_sSheetPassword).Range(C_sRngPublickey).value = T_Cle(i, 1)
    Sheets(C_sSheetPassword).Range(C_sRngPrivatekey).value = T_Cle(i, 2)

    MsgBox "My new password : " & T_Cle(i, 2)    'MSG_NewPass

    Sheets(C_sSheetPassword).Visible = xlSheetVeryHidden

End Sub

Function LetKey(bPriv As Boolean) As Long

    If bPriv Then
        LetKey = Sheets(C_sSheetPassword).Range("PrivateKey").value
    Else
        LetKey = Sheets(C_sSheetPassword).Range("PublicKey").value
    End If

End Function




