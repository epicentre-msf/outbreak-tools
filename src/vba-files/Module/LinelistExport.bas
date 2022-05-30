Attribute VB_Name = "LinelistExport"
Option Explicit

Function GetExportHeaders(sExportName As String, sSheetName As String, Optional isMigration As Boolean = False) As BetterArray
    Dim ExportHeaders As BetterArray
    Dim SheetNameData As BetterArray
    Dim YesNoExportData As BetterArray

    Dim i As Integer

    Set ExportHeaders = New BetterArray
    ExportHeaders.LowerBound = 1

    Set SheetNameData = GetDictionaryColumn(C_sDictHeaderSheetName)

    If Not isMigration Then
        Set YesNoExportData = GetDictionaryColumn(sExportName)
        For i = 1 To YesNoExportData.UpperBound
           If ClearString(YesNoExportData.Item(i)) = "yes" And SheetNameData.Item(i) = sSheetName Then
             ExportHeaders.Push ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(i + 1, 1).value 'Take the varname
           End If
        Next
    Else
        For i = 1 To SheetNameData.UpperBound
           If SheetNameData.Item(i) = sSheetName Then
             ExportHeaders.Push ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(i + 1, 1).value 'Take all the varname
           End If
        Next
    End If

    Set GetExportHeaders = ExportHeaders.Clone()
    Set YesNoExportData = Nothing
    Set SheetNameData = Nothing
    Set ExportHeaders = Nothing
End Function

Function GetExportValues(ExportHeadersData As BetterArray, sSheetName As String, Optional iDataType As Integer = 1) As BetterArray

    'iDataType = 1, then it is a linelist, if iDataType = 2 then it is an admin
    'ExportHeadersData, data of the headers with columns to export.

    Dim ExportColumn As BetterArray 'one column of a data to export
    Dim SheetVarNamesData As BetterArray 'will contains all the variables for one sheet
    Dim ExportTableData As BetterArray
    Dim tempData As BetterArray 'temporary data for variable column

    Dim i As Integer 'iterator

    Set ExportTableData = New BetterArray
    Set ExportColumn = New BetterArray
    Set SheetVarNamesData = New BetterArray
    Set tempData = New BetterArray

    ExportTableData.LowerBound = 1
    ExportColumn.LowerBound = 1
    SheetVarNamesData.LowerBound = 1

    Set SheetVarNamesData = GetDictDataFromCondition(C_sDictHeaderSheetName, sSheetName, True)

    'Getting the export table depending on the datatype (1 for Linelist, 2 for admin)
    Select Case iDataType
        Case 1
            For i = 1 To ExportHeadersData.UpperBound
                If SheetVarNamesData.Includes(ExportHeadersData.Item(i)) Then
                    With ThisWorkbook.Worksheets(sSheetName)
                            'Column of filled data
                            ExportColumn.FromExcelRange .Cells(C_eStartLinesLLData + 2, SheetVarNamesData.IndexOf(ExportHeadersData.Item(i))), DetectLastColumn:=False, DetectLastRow:=True
                            'Adding the column
                            ExportTableData.Item(i) = ExportColumn.Items
                            ExportColumn.Clear
                    End With
                End If
            Next
        Case 2 'Admin data
            ExportColumn.Clear
            tempData.Clear
            For i = 1 To ExportHeadersData.UpperBound
                If SheetVarNamesData.Includes(ExportHeadersData.Item(i)) Then
                    With ThisWorkbook.Worksheets(sSheetName)
                            'Column of filled data
                            tempData.Push .Cells(C_eStartLinesAdmData + i - 1, 3) 'values
                            ExportColumn.Push ExportHeadersData.Item(i) 'variables
                    End With
                End If
            Next
            ExportTableData.Item(1) = ExportColumn.Items
            ExportTableData.Item(2) = tempData.Items
    End Select

    On Error GoTo errTranspose

    ExportTableData.ArrayType = BA_MULTIDIMENSION
    ExportTableData.Transpose
    Set GetExportValues = ExportTableData.Clone()

    Set ExportTableData = Nothing       'Table of all the export data
    Set ExportColumn = Nothing        'one column of a data to export
    Set SheetVarNamesData = Nothing   'will contains all the variables for one sheet
    Set tempData = Nothing

    Exit Function
errTranspose:
    MsgBox "Unable to transpose Export Table", vbOKOnly + VbCritical, "ERROR"
End Function


Sub Export(iTypeExport As Byte)
    Dim DictHeaders     As BetterArray 'Headers of the dictionary
    Dim LLSheetData     As BetterArray 'Vector of all sheets of type linelist
    Dim Wkb             As Workbook
    Dim DictData        As BetterArray 'Values of the dictionary
    Dim ExportData      As BetterArray
    Dim PathData        As BetterArray 'Path to exports
    Dim VarNameData     As BetterArray
    Dim ExportHeader    As BetterArray
    Dim ChoicesData     As BetterArray
    Dim TransData       As BetterArray
    Dim AdmSheetData    As BetterArray
    Dim LLExportHeader    As BetterArray

    Dim i As Integer 'Iterator
    Dim fileformat As Byte
    Dim iWindowState As Integer

    Dim sPrevSheetName As String
    Dim sPath As String
    Dim sDirectory As String
    Dim sSheetName As String
    Dim sExportName As String

    Dim AbleToExport As Boolean

    Set DictHeaders = New BetterArray
    Set DictData = New BetterArray
    Set LLSheetData = New BetterArray
    Set AdmSheetData = New BetterArray
    Set ExportData = New BetterArray
    Set PathData = New BetterArray
    Set VarNameData = New BetterArray
    Set ExportHeader = New BetterArray

    DictHeaders.LowerBound = 1
    DictData.LowerBound = 1
    LLSheetData.LowerBound = 1
    ExportData.LowerBound = 1
    PathData.LowerBound = 1
    VarNameData.LowerBound = 1
    AdmSheetData.LowerBound = 1


    On Error GoTo exportErrHandExport

    'Varname data are all the varnames in the dictionary
    Set VarNameData = GetDictionaryColumn(C_sDictHeaderVarName)
    Set DictHeaders = GetDictionaryHeaders()
    Set DictData = GetDictionaryData()
    Set LLExportHeader = Helpers.GetHeaders(ThisWorkbook, C_sParamSheetExport, 1)

    'Path to the output
    i = LLExportHeader.IndexOf("file name")

    sPath = ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value

    AbleToExport = False

    If sPath <> "" Then
        PathData.Items = Split(sPath, "+")
        i = 1
        While i <= PathData.UpperBound
            PathData.Item(i) = Replace(Application.WorksheetFunction.Trim(PathData.Items(i)), "+", "")

            If VarNameData.Includes(PathData.Items(i)) Then
                sSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                PathData.Item(i) = ThisWorkbook.Worksheets(sSheetName).Range(PathData.Items(i)).value
            Else
                PathData.Item(i) = Replace(PathData.Items(i), Chr(34), "")
            End If
            i = i + 1
        Wend

        sPath = PathData.ToString(Separator:="_", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False) & _
                        "__" & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PublicKey").value & "__" & Format(Now, "yyyymmdd-HhNn")

        sDirectory = Helpers.LoadFolder

        If sDirectory <> "" Then
                sPath = sDirectory & Application.PathSeparator & sPath

                i = 0
                While Len(sPath) >= 255 And i < 3 'MSG_PathTooLong
                    MsgBox "The path of the export file is too long so the file name gets truncated. Please select a folder higher in the hierarchy to save the export (ex: Desktop, Downloads, Documents etc.)"
                    sDirectory = LoadFolder
                    If sDirectory <> "" Then
                        sPath = sDirectory & Application.PathSeparator & sPath
                    End If
                    i = i + 1
                Wend

                'On enregistre
                If i < 3 Then
                    AbleToExport = True
                Else
                    'Unable to export, leave the program
                    F_Export.Hide
                    Exit Sub
                End If
        End If
    End If

    'Creating the data for the exports
    On Error GoTo exportErrHandData

    If AbleToExport Then
        BeginWork xlsapp:=Application
        Application.WindowState = xlMinimized

        Set Wkb = Workbooks.Add

        i = 1
        sPrevSheetName = ""
        While i <= DictData.Length
            'Get the list of all the Sheets of type linelist
            If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetType))) = C_sDictSheetTypeLL Then
                If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))) <> sPrevSheetName Then
                    sPrevSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                    LLSheetData.Push sPrevSheetName
                End If
            End If

            'Get the list of all the sheets if type admin
            If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetType))) = C_sDictSheetTypeAdm Then
                If (DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))) <> sPrevSheetName Then
                    sPrevSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                    AdmSheetData.Push sPrevSheetName
                End If
            End If

            i = i + 1
        Wend

        With Wkb

            'Adding the worksheets
            sPrevSheetName = .Worksheets(1).Name


            'Add Translation
            i = LLExportHeader.IndexOf("export translation")

            If (ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value = "yes") Then
                 Set TransData = GetTransData()
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetTranslation
                TransData.ToExcelRange .Worksheets(C_sParamSheetTranslation).Cells(1, 1)
                 sPrevSheetName = C_sParamSheetTranslation
            End If

            i = LLExportHeader.IndexOf("export dictionary")

            'Add Choice
            If (ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value = "yes") Then

                Set ChoicesData = GetChoicesData()
                'Choices Sheet
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetChoices
                ChoicesData.ToExcelRange .Worksheets(C_sParamSheetChoices).Cells(1, 1)
                sPrevSheetName = C_sParamSheetChoices

                'Add Dictionary
                'Writing the dictionary data
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetDict
                DictHeaders.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
                DictData.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(2, 1)
                sPrevSheetName = C_sParamSheetDict
            End If

        'Adding the others sheets (Admin, linelist)
        'Get the type of the export
        Select Case iTypeExport
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

        'write all the sheets of type linelist
        i = 1
        While i <= LLSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = LLSheetData.Items(i)
            sPrevSheetName = LLSheetData.Items(i)
            ExportData.Clear
            ExportHeader.Clear

            Set ExportHeader = GetExportHeaders(sExportName, sPrevSheetName)
            Set ExportData = GetExportValues(ExportHeader, sPrevSheetName)
            ExportHeader.ToExcelRange .Worksheets(sPrevSheetName).Cells(1, 1), TransposeValues:=True
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(2, 1)
            i = i + 1
        Wend

        'write all the sheets of type admin
        i = 1
        While i <= AdmSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = AdmSheetData.Items(i)
            sPrevSheetName = AdmSheetData.Items(i)
            ExportData.Clear
            ExportHeader.Clear

            Set ExportHeader = GetExportHeaders(sExportName, sPrevSheetName)
            Set ExportData = GetExportValues(ExportHeader, sPrevSheetName, 2)
            .Worksheets(sPrevSheetName).Cells(1, 1).value = "Variable"
            .Worksheets(sPrevSheetName).Cells(1, 2).value = "Value"
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(2, 1)
            i = i + 1
        Wend
    End With
    End If

    'Now writing on the choosen directory
    On Error GoTo exportErrHandWrite

    'Handling the file format
        i = LLExportHeader.IndexOf("file format")

        Select Case ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value
        Case "xlsx"
            fileformat = xlOpenXMLStrictWorkbook
        Case "xlsb"
            fileformat = xlExcel12
        Case Else
            fileformat = xlExcel12
        End Select


    'Handing password

    i = LLExportHeader.IndexOf("password")

    Select Case ClearString(ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value)
        Case "yes"
          Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, Password:=ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value, _
          ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
          MsgBox "File saved" & Chr(10) & "Password : " & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value
        Case "no"
         Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
          MsgBox "File saved" & Chr(10) & "No Password"
        Case Else
         Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, Password:=ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value, _
         ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
          MsgBox "File saved" & Chr(10) & "Password : " & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value
    End Select

    Wkb.Close
    F_Export.Hide
    EndWork xlsapp:=Application

    Set Wkb = Nothing
    Set DictHeaders = Nothing
    Set DictData = Nothing
    Set LLSheetData = Nothing
    Set AdmSheetData = Nothing
    Set ExportData = Nothing
    Set PathData = Nothing
    Set VarNameData = Nothing
    Set ExportHeader = Nothing
    Set LLExportHeader = Nothing


    Exit Sub

exportErrHandExport:
    MsgBox "Errors during export, unable to export to corresponding path", vbOKOnly + VbCritical, "ERROR"
    Exit Sub
exportErrHandData:
    MsgBox "Errors during export, problems while getting the data", vbOKOnly + VbCritical, "ERROR"
    Exit Sub
exportErrHandWrite:
    MsgBox "Errors during export, unable to write data to corresponding directory, please choose another one", vbOKOnly + VbCritical, "ERROR"
    Exit Sub
End Sub


Sub NewKey()
    '
    Dim nbLigne As Integer
    Dim T_Cle
    Dim i As Integer

    ThisWorkbook.Worksheets(C_sSheetPassword).Visible = xlSheetHidden

    T_Cle = ThisWorkbook.Worksheets(C_sSheetPassword).ListObjects(C_sTabkeys).DataBodyRange
    nbLigne = UBound(T_Cle, 1)

   'Randomize
    i = Int(nbLigne * Rnd())
    ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngPublickey).value = T_Cle(i, 1)
    ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngPrivatekey).value = T_Cle(i, 2)

    MsgBox "My new password : " & T_Cle(i, 2)    'MSG_NewPass

    ThisWorkbook.Worksheets(C_sSheetPassword).Visible = xlSheetVeryHidden

End Sub

Function LetKey(bPriv As Boolean) As Long

    If bPriv Then
        LetKey = ThisWorkbook.Worksheets(C_sSheetPassword).Range("PrivateKey").value
    Else
        LetKey = ThisWorkbook.Worksheets(C_sSheetPassword).Range("PublicKey").value
    End If

End Function


