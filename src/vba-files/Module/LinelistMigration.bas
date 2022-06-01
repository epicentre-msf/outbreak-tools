Attribute VB_Name = "LinelistMigration"
Option Explicit



'Need to add and clean the import for the data.


'=========================== IMPORTS MIGRATIONS ================================

'Import Migration Data

Sub ImportMigrationData()

'Import exported data into the linelist

   'Import exported data into the linelist

    Dim wbkexp As Workbook, wbkLL As Workbook
    Dim shData As Worksheet, shSource As Worksheet, shTarget As Worksheet, ShMain As Worksheet
    Dim lstobj  As ListObject
    Dim iLastSh As Integer, iLastexp As Integer, i As Long, j As Long
    Dim lgRows As Long, iCols As Integer, lgStart As Long, iColExp As Integer, lgRowTarget As Long, iColTarget As Integer, lgNbData As Long
    Dim sLabel As String, sPath As String, sMessage As String
    Dim iRep As Integer
    
    Set ShMain = Sheets(ActiveSheet.Name)
    
    Set wbkLL = ActiveWorkbook

    iLastSh = wbkLL.Sheets.Count
    iLastexp = iLastSh

    sPath = LoadFile("*.xlsx, *.xlsb")

    If sPath = "" Then Exit Sub

    For Each shData In wbkLL.Sheets
        If shData.Visible = xlSheetVisible And LCase(shData.Name) <> "geo" And LCase(shData.Name) <> "admin" Then
            For Each lstobj In shData.ListObjects
                If LCase(lstobj.Name) = "o" & LCase(shData.Name) Then
                    lgRowTarget = shData.ListObjects(lstobj.Name).ListRows.Count + 8
                    iColTarget = shData.ListObjects(lstobj.Name).ListColumns.Count
                    Application.ScreenUpdating = False
                        shData.Activate
                            lgNbData = WorksheetFunction.CountA(shData.Range(Cells(8, 1), Cells(lgRowTarget, iColTarget)))
                        ShMain.Activate
                    Application.ScreenUpdating = True
                    If lgNbData > 0 Then
                        If iRep = 0 Then iRep = MsgBox("There is data in the sheets. Want to keep them or delete them ?", vbQuestion + vbYesNo, "Import Data")
                        If iRep = vbYes Then
                            Application.EnableEvents = False
                                shData.Unprotect (C_sLLPassword)
                                    lstobj.DataBodyRange.Delete
                                Call ProtectSheet
                            Application.EnableEvents = True
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    If iRep = vbYes Then
        iRep = 0
        lgRows = ShMain.Cells(15, 2).End(xlDown).Row
        ShMain.Unprotect (C_sLLPassword)
            ShMain.Range(Cells(15, 3), Cells(lgRows, 3)).ClearContents
        Call ProtectSheet
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Workbooks.Open Filename:=sPath, Password:=Range("RNG_PrivateKey").value

    Set wbkexp = ActiveWorkbook

    For Each shData In wbkexp.Sheets

        If LCase(shData.Name) <> "dictionary" And LCase(shData.Name) <> "choices" And LCase(shData.Name) <> "translation" Then

            For i = 1 To wbkLL.Sheets.Count

                If UCase(wbkLL.Sheets(i).Name) = UCase(shData.Name) Then
                    Sheets(shData.Name).Copy After:=wbkLL.Sheets(iLastexp)
                    ActiveSheet.Name = shData.Name & "_Exp"
                    iLastexp = iLastexp + 1
                    wbkexp.Activate
                    Exit For
                End If

            Next i

        End If

    Next shData

    wbkexp.Close

    Set wbkexp = Nothing
    Set wbkLL = Nothing

    For i = iLastexp To iLastSh + 1 Step -1

        Set shSource = Sheets(Sheets(i).Name)
        Set shTarget = Sheets(Left(Sheets(i).Name, Len(Sheets(i).Name) - 4))

        If LCase(shSource.Name) = "admin_exp" Then
        
            shTarget.Unprotect (C_sLLPassword)
            shSource.Select
            lgRows = shSource.Cells(2, 1).End(xlDown).Row
            
            For j = 2 To lgRows
                Application.EnableEvents = False
                    If shTarget.Cells(j + 13, 3).value = "" Then shSource.Cells(j, 2).Copy Destination:=shTarget.Cells(j + 13, 3)
                Application.EnableEvents = True
            Next j

        Else

            iCols = shSource.Cells(1, 1).End(xlToRight).Column
            lgRows = shSource.Cells(1, 1).End(xlDown).Row
            lgStart = shTarget.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1

            shTarget.Select
            ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

            For Each lstobj In shTarget.ListObjects
                If lstobj.Name = "o" & shTarget.Name Then
                    lgRowTarget = shTarget.ListObjects(lstobj.Name).ListRows.Count
                    If lgRowTarget < lgRows Then
                        Application.EnableEvents = False
                        lstobj.Resize Range(Cells(lgStart - 1, 1), Cells(lgRows + lgStart, Cells(lgStart - 1, 1).End(xlToRight).Column))
                        Application.EnableEvents = True
                        lgRowTarget = shTarget.ListObjects(lstobj.Name).ListRows.Count
                        Exit For
                    End If
                End If
            Next

            ShMain.Select

            j = 1


            Do While j <= iCols 'Number of columns in source file

                If shTarget.Cells(6, j).value = "" Then Exit Do
                Dim BlHidden As Boolean

                If shTarget.Columns(j).Hidden = True Then
                    shTarget.Columns(j).Hidden = False
                    BlHidden = True
                End If

                sLabel = shSource.Cells(1, j).value

                If Not shTarget.Rows(7).Find(What:=sLabel, LookAt:=xlWhole) Is Nothing Then

                    iColExp = shSource.Rows(1).Find(What:=sLabel, LookAt:=xlWhole).Column

                    shSource.Select

                    iColTarget = shTarget.Rows(7).Find(What:=sLabel, LookAt:=xlWhole).Column

                    If iColTarget > 0 And Not shTarget.Cells(8, j).HasFormula Then
                        Application.EnableEvents = False
                            shSource.Range(Cells(2, iColExp), Cells(lgRows, iColExp)).Copy Destination:=shTarget.Cells(lgStart, iColTarget)
                        Application.EnableEvents = True
                        shTarget.Columns(iColTarget).EntireColumn.AutoFit
                    End If
                    
                Else

                    sMessage = sMessage & " / " & sLabel & " (Sheet " & Replace(shSource.Name, "_Exp", "") & ")"
                    
                End If

                If BlHidden Then
                    shTarget.Columns(j).Hidden = True
                    BlHidden = False
                End If

                j = j + 1
            Loop

        End If

        Application.DisplayAlerts = False
            shSource.Delete
        Application.DisplayAlerts = True

        Call ProtectSheet

    Next i

    ShMain.Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "the following data :" & Chr(10) & Right(sMessage, Len(sMessage) - 3) & Chr(10) & "could not be imported.", vbInformation, "File import"

End Sub
'Function to import a Geobase.
'Import the full Geobase

'Function to import a Geobase.
'Import the full Geobase

Sub ImportGeobase()

    BeginWork xlsapp:=Application

    Dim sFilePath   As String                      'File path to the geo file
    Dim oSheet      As Object
    Dim AdmData     As BetterArray                  'Table for admin levels
    Dim AdmHeader   As BetterArray                 'Table for the headers of the listobjects
    Dim AdmNames    As BetterArray                  'Array of the sheetnames
    Dim i           As Integer                             'iterator
    Dim Wkb         As Workbook
    Dim WkshGeo     As Worksheet
    Dim ShouldQuit As Integer
    'Sheet names
    Set AdmNames = New BetterArray
    Set AdmData = New BetterArray
    Set AdmHeader = New BetterArray

    AdmNames.LowerBound = 1
    AdmNames.Push C_sAdm1, C_sAdm2, C_sAdm3, C_sAdm4, C_sHF, C_sNames, C_sHistoHF, C_sHistoGeo, C_sGeoMetadata  'Names of each sheet

    'Set xlsapp = New Excel.Application
    sFilePath = Helpers.LoadFile("*.xlsx")

    If sFilePath <> "" Then
        'Open the geo workbook and hide the windows
        Set Wkb = Workbooks.Open(sFilePath)
        Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

        'Write the filename of the geobase somewhere for the export
        WkshGeo.Range(C_sRngGeoName).value = Dir(sFilePath)

        For i = 1 To AdmNames.Length
            'Adms (Maybe come back to work on the names?)
            If Not WkshGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange.Delete
            End If
        Next

        'Reloading the data from the Geobase
        For Each oSheet In Wkb.Worksheets
            AdmData.Clear
            AdmHeader.Clear

            'Be sure my sheetnames are correct before loading the data
            If AdmNames.Includes(oSheet.Name) Then

                'loading the data in memory
                AdmData.FromExcelRange oSheet.Range("A2"), DetectLastRow:=True, DetectLastColumn:=True
                'The headers
                AdmHeader.FromExcelRange oSheet.Range("A1"), DetectLastRow:=False, DetectLastColumn:=True

                'Check if the sheet is the admin exists sheet before writing in the adm table
                With WkshGeo.ListObjects("T_" & oSheet.Name)
                    AdmHeader.ToExcelRange Destination:=WkshGeo.Cells(1, .Range.Column), TransposeValues:=True
                    AdmData.ToExcelRange Destination:=WkshGeo.Cells(2, .Range.Column)

                    'Resizing the Table
                    .Resize .Range.CurrentRegion
                End With

            End If
        Next


        Wkb.Close savechanges:=False


        Call TranslateImportGeoHead

        ShouldQuit = MsgBox("Finished imported geobase data, do you want to quit?", vbQuestion + vbYesNo, "Import GeoData")

        If ShouldQuit = vbYes Then
            F_ImportMig.Hide
        End If
    End If

    Set AdmHeader = Nothing
    Set AdmNames = Nothing
    Set Wkb = Nothing
    Set AdmData = Nothing
    Set Wkb = Nothing
    Set WkshGeo = Nothing

    EndWork xlsapp:=Application
End Sub


Private Sub TranslateImportGeoHead()

    Dim sIsoCountry As String
    Dim sCountry As String
    Dim sSubCounty As String
    Dim sWard As String
    Dim sPlace As String
    Dim sFacility As String
    Dim Wksh As Worksheet

    Set Wksh = ThisWorkbook.Worksheets(C_sSheetLLTranslation)

    'Get the isoCode for the linelist
    sIsoCountry = Wksh.Range(C_sRngLLLanguageCode).value
    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)

    sCountry = Application.WorksheetFunction.HLookup(sIsoCountry, Wksh.ListObjects(C_sTabNames).Range, 2, False)
    sSubCounty = Application.WorksheetFunction.HLookup(sIsoCountry, Wksh.ListObjects(C_sTabNames).Range, 3, False)
    sWard = Application.WorksheetFunction.HLookup(sIsoCountry, Wksh.ListObjects(C_sTabNames).Range, 4, False)
    sPlace = Application.WorksheetFunction.HLookup(sIsoCountry, Wksh.ListObjects(C_sTabNames).Range, 5, False)
    sFacility = Application.WorksheetFunction.HLookup(sIsoCountry, Wksh.ListObjects(C_sTabNames).Range, 6, False)

    Wksh.Range("A1,E1,J1,P1,Z1").value = sCountry
    Wksh.Range("F1,K1,Q1,Y1").value = sSubCounty
    Wksh.Range("L1,R1,X1").value = sWard
    Wksh.Range("S1").value = sPlace
    Wksh.Range("W1").value = sFacility

End Sub


'Import the only the Historic Data
Sub ImportHistoricGeobase()

    On Error GoTo errImportHistoric

    BeginWork xlsapp:=Application

    Dim sFilePath   As String                      'File path to the geo file
    Dim oSheet      As Object
    Dim AdmData     As BetterArray                  'Table for admin levels
    Dim AdmNames    As BetterArray                  'Array of the sheetnames
    Dim i           As Integer                             'iterator
    Dim Wkb         As Workbook
    Dim WkshGeo     As Worksheet
    Dim ShouldQuit As Integer

    'Sheet names
    Set AdmNames = New BetterArray
    Set AdmData = New BetterArray

    AdmNames.LowerBound = 1
    AdmNames.Push C_sHistoHF, C_sHistoGeo, C_sGeoMetadata 'Names of each sheet

    sFilePath = Helpers.LoadFile("*.xlsx")

    If sFilePath <> "" Then
        'Open the geo workbook and hide the windows
        Set Wkb = Workbooks.Open(sFilePath)
        Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

        For i = 1 To AdmNames.Length
            'Adms (Maybe come back to work on the names?)
            If Not WkshGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange.Delete
            End If
        Next

        'Reloading the data from the Geobase
        For Each oSheet In Wkb.Worksheets
            AdmData.Clear

            'Be sure my sheetnames are correct before loading the data
            If AdmNames.Includes(oSheet.Name) Then

                'loading the data in memory
                AdmData.FromExcelRange oSheet.Range("A2"), DetectLastRow:=True, DetectLastColumn:=False

                'Check if the sheet is the admin exists sheet before writing in the adm table
                With WkshGeo.ListObjects("T_" & oSheet.Name)
                    AdmData.ToExcelRange Destination:=WkshGeo.Cells(2, .Range.Column)
                    'Resizing the Table
                    .Resize .Range.CurrentRegion
                End With

            End If
        Next
        Wkb.Close savechanges:=False
        'Add a message box to say it is over
    End If

    ShouldQuit = MsgBox("Finished Imported historic data, do you want to quit?", vbQuestion + vbYesNo, "Import Historic")

    If ShouldQuit = vbYes Then
        F_ImportMig.Hide
    End If

    Set AdmData = Nothing
    Set AdmNames = Nothing
    Set Wkb = Nothing
    Set WkshGeo = Nothing

    EndWork xlsapp:=Application

    Exit Sub

errImportHistoric:
        MsgBox "Error, Unable to import historic geobase"
        EndWork xlsapp:=Application
        Exit Sub

End Sub

'Clear the historic Data

Sub ClearHistoricGeobase()
    Dim WkshGeo As Worksheet
    Dim ShouldDelete As Integer

    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    On Error GoTo errClearHistoric

    ShouldDelete = MsgBox("Your historic geographic data  in the current workbook will be completely deleted, this action is irreversible. Proceed?", vbExclamation + vbYesNo, "Delete Historic")

    If ShouldDelete = vbYes Then
        If Not WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
            WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange.Delete
        End If

        If Not WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
            WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange.Delete
        End If
        MsgBox "Done", vbInformation, "Delete Historic"
    End If

    'Add a message to say it is done
    Set WkshGeo = Nothing

    Exit Sub

errClearHistoric:
        MsgBox "Error, Unable to clear the historic of geobase"
        EndWork xlsapp:=Application
        Exit Sub
End Sub



'============================= EXPORTS MIGRATIONS ==============================


'Export the data

Private Sub ExportMigrationData(sLLPath As String)


    'Dictionary headers and data
    Dim DictHeaders As BetterArray
    Dim DictData As BetterArray

    'Linelist and Admin Sheets data

    Dim LLSheetData As BetterArray
    Dim AdmSheetData As BetterArray

    'Export data and headers
    Dim ExportData As BetterArray
    Dim ExportHeader As BetterArray

    Dim i As Integer 'iterator
    Dim sPrevSheetName As String 'Previous sheet name  used as temporary variable

    Dim Wkb As Workbook

    Set LLSheetData = New BetterArray
    Set AdmSheetData = New BetterArray
    Set ExportData = New BetterArray
    Set ExportHeader = New BetterArray

    On Error GoTo errExportMig

    LLSheetData.LowerBound = 1
    AdmSheetData.LowerBound = 1

    'Now I am able to Export, try to write each data to a workbook
    BeginWork xlsapp:=Application
    Application.DisplayAlerts = False

    'Initialize ditionary headers and values
    Set DictHeaders = GetDictionaryHeaders()
    Set DictData = GetDictionaryData()

    i = 1
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

    'Writing the linelist Data (with all the databases, dictionary, export, translation and choices)
    Set Wkb = Workbooks.Add

    With Wkb
        'Writing the translation data
        sPrevSheetName = .Worksheets(1).Name
        Set ExportData = GetTransData()
        .Worksheets(sPrevSheetName).Name = C_sParamSheetTranslation
        ExportData.ToExcelRange .Worksheets(C_sParamSheetTranslation).Cells(1, 1)
        sPrevSheetName = C_sParamSheetTranslation

        'Writing the choice data
        Set ExportData = GetChoicesData()
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetChoices
        ExportData.ToExcelRange .Worksheets(C_sParamSheetChoices).Cells(1, 1)
        sPrevSheetName = C_sParamSheetChoices

        'Writing the dictionary
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetDict
        DictHeaders.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
        DictData.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(2, 1)
        sPrevSheetName = C_sParamSheetDict

        'Sheets of type linelist
        i = 1
        While i <= LLSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = ClearString(LLSheetData.Items(i), bremoveHiphen:=False)
            sPrevSheetName = ClearString(LLSheetData.Items(i), bremoveHiphen:=False)
            ExportData.Clear
            ExportHeader.Clear
            Set ExportHeader = GetExportHeaders("Migration", sPrevSheetName, isMigration:=True)
            Set ExportData = GetExportValues(ExportHeader, sPrevSheetName)
            ExportHeader.ToExcelRange .Worksheets(sPrevSheetName).Cells(1, 1), TransposeValues:=True
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(2, 1)
            i = i + 1
        Wend

        'Sheets of type Admin
        i = 1
        While i <= AdmSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = ClearString(AdmSheetData.Items(i), bremoveHiphen:=False)
            sPrevSheetName = ClearString(AdmSheetData.Items(i), bremoveHiphen:=False)
            ExportData.Clear
            ExportHeader.Clear
            Set ExportHeader = GetExportHeaders("Migration", sPrevSheetName, isMigration:=True)
            Set ExportData = GetExportValues(ExportHeader, sPrevSheetName, 2)
            .Worksheets(sPrevSheetName).Cells(1, 1).value = "variable"
            .Worksheets(sPrevSheetName).Cells(1, 2).value = "value"
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(2, 1)
            i = i + 1
        Wend

        'Add The metadata Sheet

         .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sSheetMetadata
         sPrevSheetName = C_sSheetMetadata
        .Worksheets(sPrevSheetName).Cells(1, 1).value = "variable"
        .Worksheets(sPrevSheetName).Cells(1, 2).value = "value"
        'Writing informations for metadata
        .Worksheets(sPrevSheetName).Cells(2, 1).value = "language"
        .Worksheets(sPrevSheetName).Cells(2, 2).value = ThisWorkbook.Worksheets(C_sSheetLLTranslation).Range(C_sRngLLLanguage)
        'Will add other metadata in the future
    End With

    'Write an error handling for writing the file here
    Wkb.SaveAs Filename:=sLLPath, fileformat:=xlExcel12, CreateBackup:=False, _
     ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Wkb.Close

    Set Wkb = Nothing
    Set ExportData = Nothing
    Set ExportHeader = Nothing
    Set LLSheetData = Nothing
    Set AdmSheetData = Nothing
    Set DictData = Nothing
    Set DictHeaders = Nothing

    Application.DisplayAlerts = True
    EndWork xlsapp:=Application

    Exit Sub

errExportMig:
        MsgBox "Error, unable to export data"
        EndWork xlsapp:=Application
        Exit Sub
End Sub

'Export the Full Geobase (including historic)


Private Sub ExportMigrationGeo(sGeoPath As String)

    Dim Wkb As Workbook
    Dim WkshGeo As Worksheet
    Dim ExportData As BetterArray
    Dim ExportHeader As BetterArray
    Dim sPrevSheetName As String


    BeginWork xlsapp:=Application
    Application.DisplayAlerts = False

    On Error GoTo errExportMigGeo

    Set Wkb = Workbooks.Add
    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    Set ExportData = New BetterArray
    Set ExportHeader = New BetterArray
    ExportHeader.LowerBound = 1
    ExportData.LowerBound = 1

    With Wkb
        sPrevSheetName = .Worksheets(1).Name
        'Add the worksheets for each of the ADM and Histo Levels

        'Histo HF
        .Worksheets(sPrevSheetName).Name = C_sHistoHF
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabHistoHF).Range
        ExportData.ToExcelRange .Worksheets(C_sHistoHF).Cells(1, 1)
        sPrevSheetName = C_sHistoHF

        'Histo Geo
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sHistoGeo
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabHistoGeo).Range
        ExportData.ToExcelRange .Worksheets(C_sHistoGeo).Cells(1, 1)
        sPrevSheetName = C_sHistoGeo
        ExportData.Clear

        'NAMES
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sNames
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabNames).Range
        ExportData.ToExcelRange .Worksheets(C_sNames).Cells(1, 1)
        sPrevSheetName = C_sNames
        ExportData.Clear

        'HF
        ExportHeader.Clear
        ExportHeader.Push LCase(C_sHF) & "_name", LCase(C_sAdm3) & "_name", LCase(C_sAdm2) & "_name", LCase(C_sAdm1) & "_name"
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sHF
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabHF).Range
        ExportData.ToExcelRange .Worksheets(C_sHF).Cells(1, 1)
        ExportHeader.ToExcelRange .Worksheets(C_sHF).Cells(1, 1), TransposeValues:=True
        sPrevSheetName = C_sHF
        ExportData.Clear

        'I need the same headers as the geobase for the admin part
        ExportHeader.Clear
        ExportHeader.Push LCase(C_sAdm1) & "_name", LCase(C_sAdm2) & "_name", LCase(C_sAdm3) & "_name", LCase(C_sAdm4) & "_name"
        'ADM4
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sAdm4
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabAdm4).Range
        ExportData.ToExcelRange .Worksheets(C_sAdm4).Cells(1, 1)
        ExportHeader.ToExcelRange .Worksheets(C_sAdm4).Cells(1, 1), TransposeValues:=True
        .Worksheets(C_sAdm4).Cells(1, 5).value = LCase(C_sAdm4) & "_pop"
        sPrevSheetName = C_sAdm4
        ExportData.Clear

        'ADM3
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sAdm3
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabAdm3).Range
        ExportHeader.Pop
        ExportData.ToExcelRange .Worksheets(C_sAdm3).Cells(1, 1)
        ExportHeader.ToExcelRange .Worksheets(C_sAdm3).Cells(1, 1), TransposeValues:=True
        sPrevSheetName = C_sAdm3
        ExportData.Clear

        'ADM2
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sAdm2
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabAdm2).Range
        ExportData.ToExcelRange .Worksheets(C_sAdm2).Cells(1, 1)
        ExportHeader.Pop
        ExportHeader.ToExcelRange .Worksheets(C_sAdm2).Cells(1, 1), TransposeValues:=True
        sPrevSheetName = C_sAdm2
        ExportData.Clear

        'ADM1
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sAdm1
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabadm1).Range
        ExportData.ToExcelRange .Worksheets(C_sAdm1).Cells(1, 1)
        .Worksheets(C_sAdm1).Cells(1, 1).value = C_sAdmName & "1" & "_name"
        sPrevSheetName = C_sAdm1

        'Metadata
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sGeoMetadata
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabGeoMetadata).Range
        ExportData.ToExcelRange .Worksheets(C_sGeoMetadata).Cells(1, 1)

    End With

    'Writing the Geo
    Wkb.SaveAs Filename:=sGeoPath, fileformat:=xlOpenXMLStrictWorkbook, _
               CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Wkb.Close

    Application.DisplayAlerts = True
    EndWork xlsapp:=Application

    Exit Sub

errExportMigGeo:
        MsgBox "Error, unable to Export Geobase"
        EndWork xlsapp:=Application
        Exit Sub
End Sub


'Export the Geobase Historic Only

Private Sub ExportMigrationHistoricGeo(sGeoPath As String)

    Dim Wkb As Workbook
    Dim WkshGeo As Worksheet
    Dim ExportData As BetterArray


    Dim sPrevSheetName As String

    On Error GoTo ErrExportHistGeo

    BeginWork xlsapp:=Application
    Application.DisplayAlerts = False

    Set Wkb = Workbooks.Add
    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    Set ExportData = New BetterArray


    With Wkb
        sPrevSheetName = .Worksheets(1).Name
        'Add the worksheets for each of the ADM and Histo Levels

        'Histo HF
        .Worksheets(sPrevSheetName).Name = C_sHistoHF
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabHistoHF).Range
        ExportData.ToExcelRange .Worksheets(C_sHistoHF).Cells(1, 1)
        'Add headers // Remember to use the convention with the headers defined previously
        .Worksheets(C_sHistoHF).Cells(1, 1).value = C_sHistoHF
        sPrevSheetName = C_sHistoHF
        ExportData.Clear

        'Histo Geo
        .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sHistoGeo
        ExportData.FromExcelRange WkshGeo.ListObjects(C_sTabHistoGeo).Range
        ExportData.ToExcelRange .Worksheets(C_sHistoGeo).Cells(1, 1)
        'Change the headers
        .Worksheets(C_sHistoGeo).Cells(1, 1).value = C_sHistoGeo

    End With

    'Writing the Geo
    Wkb.SaveAs Filename:=sGeoPath, fileformat:=xlOpenXMLStrictWorkbook, _
               CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Wkb.Close

    Set Wkb = Nothing
    Set WkshGeo = Nothing
    Set ExportData = Nothing

    Application.DisplayAlerts = True
    EndWork xlsapp:=Application

    Exit Sub

ErrExportHistGeo:
        MsgBox "Error, unable to export the historic of the geobase"
        EndWork xlsapp:=Application
        Exit Sub
End Sub


'Export for Migration (Either historic, the full geobase and / or data)

Sub ExportForMigration()

    Dim ExportPath As BetterArray

    'boolean for controling the export
    Dim AbleToExport As Boolean

    Dim sDirectory As String 'Folder for export
    Dim sLLPath As String 'Linelist file
    Dim sGeoPath As String 'Geo File
    Dim sGeoHistoPath As String 'Historic File
    Dim sPath As String
    Dim ShouldQuit As Byte

    Dim i As Integer 'iterator

    Set ExportPath = New BetterArray

    'Select the Folder
    AbleToExport = False
    sDirectory = Helpers.LoadFolder

    On Error GoTo ErrPath

    If sDirectory <> "" Then
        'Export the Data of the linelist
        sLLPath = sDirectory & Application.PathSeparator & Replace(ClearString(ThisWorkbook.Name), ".xlsb", "") & _
             "_export_data"

        'Export the full geobase

        'For the GeoPath we need more involve work
        ExportPath.Items = Split(ThisWorkbook.Worksheets(C_sSheetGeo).Range(C_sRngGeoName).value, "_")
        ExportPath.Pop
        sPath = ExportPath.ToString(Separator:="_", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
        sGeoPath = sDirectory & Application.PathSeparator & sPath & Format(Now, "yyyymmdd")
        sGeoHistoPath = sDirectory & Application.PathSeparator & sPath & Format(Now, "yyyymmdd") & "_historic"

        i = 0
        While Len(sLLPath) >= 255 And Len(sGeoPath) >= 255 And Len(sGeoHistoPath) >= 255 And i < 3 'MSG_PathTooLong
            MsgBox "The path of the export folder is too long so the file names gets truncated. Please select a folder higher in the hierarchy to save the export (ex: Desktop, Downloads, Documents etc.)"
            sDirectory = Helpers.LoadFolder
            If sDirectory <> "" Then
                 sLLPath = sDirectory & Application.PathSeparator & Replace(ClearString(ThisWorkbook.Name), ".xlsb", "") & _
                    "_export_data"

                    'Export the full geobase

                    'For the GeoPath we need more involve work
                    ExportPath.Items = Split(ThisWorkbook.Worksheets(C_sSheetGeo).Range(C_sRngGeoName).value, "_")
                    ExportPath.Pop
                    sPath = ExportPath.ToString(Separator:="_", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    sGeoPath = sDirectory & Application.PathSeparator & sPath & Format(Now, "yyyymmdd")
                    sGeoHistoPath = sDirectory & Application.PathSeparator & sPath & Format(Now, "yyyymmdd") & "_historic"
            End If
            i = i + 1
        Wend
        If i < 3 Then
           AbleToExport = True
        Else
        'Unable to export, leave the program
         F_ExportMig.Hide
        Exit Sub
        End If
    End If

    On Error GoTo 0

    'Add here error handling when the export is not working.

    If AbleToExport Then
        'Now I am able to Export, try to write each data to a workbook
        BeginWork xlsapp:=Application
        Application.DisplayAlerts = False

        If Not [F_ExportMig].CHK_ExportMigGeo And Not [F_ExportMig].CHK_ExportMigData And Not [F_ExportMig].CHK_ExportMigGeoHistoric Then
            MsgBox "No Export Selected, Please selectOne", vbCritical, "Export for Migration"
        End If

        If [F_ExportMig].CHK_ExportMigGeo Then
            Call ExportMigrationGeo(sGeoPath)
        End If

        If [F_ExportMig].CHK_ExportMigData Then
            Call ExportMigrationData(sLLPath)
        End If

        If [F_ExportMig].CHK_ExportMigGeoHistoric Then
            Call ExportMigrationHistoricGeo(sGeoHistoPath)
        End If

        'Return previous sate of the application
        Application.DisplayAlerts = True
        EndWork xlsapp:=Application

        ShouldQuit = MsgBox("Finished Exports, do you want to quit?", vbQuestion + vbYesNo, "Export for Migration")
        If ShouldQuit = vbYes Then
            F_ExportMig.Hide
        End If

    End If

    Set ExportPath = Nothing

    Exit Sub
ErrPath:
        MsgBox "Error, unable to set the paths for the export"
        EndWork xlsapp:=Application
        Exit Sub
End Sub

Public Function SheetExist(SheetName As String) As Boolean
'check sheet exist

    Dim shSheet As Variant
    
    SheetExist = False
    
    For Each shSheet In Sheets
        If UCase(shSheet.Name) = UCase(SheetName) Then
            SheetExist = True
            Exit Function
        End If
    Next shSheet
    
End Function



