Attribute VB_Name = "LinelistMigration"
Option Explicit

Dim ImportReport As Boolean
Dim ImportVarData As BetterArray

Sub ControlClearData()

    Dim ShouldProceed As Byte
    Dim sLLName As String

    ShouldProceed = MsgBox(TranslateLLMsg("MSG_DeleteAllData"), vbExclamation + vbYesNo, TranslateLLMsg("MSG_Delete"))

    If ShouldProceed = vbYes Then
        sLLName = InputBox(TranslateLLMsg("MSG_LLName"), TranslateLLMsg("MSG_Delete"), TranslateLLMsg("MSG_EnterWkbName"))
        If sLLName = Replace(ThisWorkbook.Name, "." & GetFileExtension(ThisWorkbook.Name), "") Then
            'Proceed only if the user is able to provide the name of the actual workbook name, we can delete
            Call ClearData
        Else
            MsgBox TranslateLLMsg("MSG_BadLLName"), vbOKOnly, TranslateLLMsg("MSG_Delete")
            Exit Sub
        End If

        ShouldProceed = MsgBox(TranslateLLMsg("MSG_FinishedClear"), vbQuestion + vbYesNo, TranslateLLMsg("MSG_Imports"))
        If ShouldProceed = vbYes Then
            [F_ImportMig].Hide
        End If
    Else
        MsgBox TranslateLLMsg("MSG_DelCancel"), vbOKOnly, TranslateLLMsg("MSG_Delete")
        Exit Sub
    End If

End Sub


'Clear all the Linelist and Adm Data

Sub ClearData()

    Dim Wksh As Worksheet
    Dim sSheetType As String
    Dim iLastRow As Integer
    Dim i As Integer

    BeginWork xlsapp:=Application
    Application.EnableEvents = False

    For Each Wksh In ThisWorkbook.Worksheets
        sSheetType = FindSheetType(Wksh.Name)
        Select Case sSheetType

            Case C_sDictSheetTypeLL
                 Wksh.Unprotect ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value

                With Wksh.ListObjects("o" & ClearString(Wksh.Name))

                    If Not .DataBodyRange Is Nothing Then
                        'Delete the data body range
                        .DataBodyRange.Delete
                    End If
                End With

                Wksh.Protect Password:=ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value, _
                     DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                     AllowFormattingColumns:=True

            Case C_sDictSheetTypeAdm
                'Find Last row of Adm Data
                With Wksh
                    iLastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
                    For i = C_eStartLinesAdmData To iLastRow
                        .Cells(i, 3).value = vbNullString
                    Next
                End With

            Case Else

        End Select
    Next

    ThisWorkbook.save
    EndWork xlsapp:=Application
    Application.EnableEvents = True

End Sub

'=========================== IMPORTS MIGRATIONS ================================

Function LLhasData() As Boolean

    Dim hasData As Boolean
    Dim shLL As Worksheet
    Dim WkbLL As Workbook
    Dim TabLL As BetterArray
    Dim ShouldProceed As Byte
    Dim sLLName As String
    Dim NotGood As Boolean

    hasData = False

    Set TabLL = New BetterArray
    Set WkbLL = ThisWorkbook

    For Each shLL In WkbLL.Worksheets
        If FindSheetType(shLL.Name) = C_sDictSheetTypeLL Then

            If FindLastRow(shLL) > C_eStartLinesLLData + 2 Then
                hasData = True
                Exit For
            End If

        End If
    Next

    If hasData Then
        'Add message to as the user if he wants to delete
        ShouldProceed = MsgBox(TranslateLLMsg("MSG_DeleteForImport"), vbExclamation + vbYesNo, TranslateLLMsg("MSG_Delete"))
        If ShouldProceed = vbYes Then
            'A simple logic to test if the linelist name is not good
            NotGood = True
            Do While NotGood
                sLLName = InputBox(TranslateLLMsg("MSG_LLName"), TranslateLLMsg("MSG_Delete"), TranslateLLMsg("MSG_EnterWkbName"))
                If sLLName = Replace(ThisWorkbook.Name, "." & GetFileExtension(ThisWorkbook.Name), "") Then
                    'Proceed only if the user is able to provide the name of the actual workbook name, we can delete
                    Call ClearData
                    hasData = False
                    NotGood = False
                Else
                    ShouldProceed = MsgBox(TranslateLLMsg("MSG_BadLLNameQ"), vbExclamation + vbYesNo, TranslateLLMsg("MSG_Delete"))

                    If ShouldProceed = vbNo Then NotGood = False

                End If
            Loop
        End If
    End If


    'If you were not able to clear the data, Set it to table lenth
    LLhasData = hasData

    Set TabLL = Nothing
    Set shLL = Nothing
    Set WkbLL = Nothing
End Function

'Import Data between two sheets:
'Check if we have to take in account data in the sheet where we have to import
'Check if a variable in the sheet to import is present in the sheet Where we have
'to paste data, if it is the case
'paste either at the end of previous data, or at the top
'If not, just skeep
'Two different procedures for sheets of type Linelist and those of type Adm


Sub ImportSheetData(sSheetName As String, shImp As Worksheet, hasData As Boolean, ColumnIndexData As BetterArray, VarNamesData As BetterArray)

    Dim sSheetType As String 'Sheet type (different procedures will apply)
    Dim iLastRowImp As Integer 'LastRow (when needed for import sheet)
    Dim iLastColImp As Integer 'Last Column for Import data of Type LL
    Dim WkbLL As Workbook
    Dim i As Integer 'Counter, for All variables
    Dim k As Integer 'Counter, unfound variables
    Dim iRowIndex As Integer 'Row index for sheets of type Adm
    Dim iColIndex As Integer 'Col index for sheets of type LL
    Dim iLastRow As Integer

    Dim sVal As String

    Dim rngImp As Range 'Range in the Import sheet
    Dim rngLL As Range 'Range in the LL sheet

    'First, sheet of type Adm

    Set WkbLL = ThisWorkbook 'The workbook of Linelist
    sSheetType = FindSheetType(sSheetName)

    With WkbLL.Worksheets(sSheetName)

        Select Case sSheetType
            'Import Data of Type Adm (No choice, we have to delete previous values)

            Case C_sDictSheetTypeAdm
                iLastRowImp = shImp.Cells(shImp.Rows.Count, 1).End(xlUp).Row

                For i = 2 To iLastRowImp '2 because the first row is for headers
                    sVal = shImp.Cells(i, 1)

                    ImportVarData.Push sVal
                    If VarNamesData.Includes(sVal) Then
                        'Get the row Index here
                        iRowIndex = ColumnIndexData.Items(VarNamesData.IndexOf(sVal))
                        .Cells(iRowIndex, 3).value = shImp.Cells(i, 2).value 'On sheets of type Adm, the third column contains values
                    Else
                        '
                    End If
                Next

            Case C_sDictSheetTypeLL

            'Import Data on a Sheet of Type LL

            'First, un protect the sheet were we need to paste the data before proceeding
            .Unprotect WkbLL.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value

            'Last column of Import Data
            iLastColImp = shImp.Cells(1, shImp.Columns.Count).End(xlToLeft).Column
            iLastRow = C_eStartLinesLLData + 2

            If hasData Then
                iLastRow = FindLastRow(WkbLL.Worksheets(sSheetName))
            End If

            k = 1
            For i = 1 To iLastColImp
                sVal = shImp.Cells(1, i)
                ImportVarData.Push sVal

                If VarNamesData.Includes(sVal) Then
                    'The variable is in the sheet, just paste the values to last row
                    iColIndex = ColumnIndexData.Items(VarNamesData.IndexOf(sVal))

                    If .Cells(C_eStartLinesLLMainSec - 1, iColIndex).value <> C_sDictControlForm Then
                        'Don't Import columns of Type formulas
                        With shImp
                            iLastRowImp = .Cells(.Rows.Count, i).End(xlUp).Row
                            Set rngImp = .Range(.Cells(2, i), .Cells(iLastRowImp, i))  '2 because first row if for headers ie varnames
                        End With
                        'Take one because of the headers in the import Sheet
                        iLastRowImp = iLastRowImp - 1
                        'Update only if there is data
                        If iLastRowImp > 0 Then
                            Set rngLL = .Range(.Cells(iLastRow, iColIndex), .Cells(iLastRow + iLastRowImp - 1, iColIndex))
                            'Copy by values, more safe
                            rngLL.value = rngImp.value
                        End If
                    End If
                Else
                    If Not ImportReport Then ImportReport = True
                    With ThisWorkbook.Worksheets(C_sSheetImportTemp)

                         k = .Cells(.Rows.Count, 3).End(xlUp).Row + 1

                        .Cells(k, 3).value = sVal
                        .Cells(k, 4).value = sSheetName
                    End With
                End If
            Next

            .Protect Password:=WkbLL.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value, _
                     DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                     AllowFormattingColumns:=True
        End Select
    End With
    Set rngImp = Nothing
End Sub

Sub ImportMigrationData()

    Dim WkbImp As Workbook 'Workbook Import

    Dim shImp As Worksheet 'Sheet Import
    Dim TabSheetLL As BetterArray 'List of sheets in the linelist
    Dim VarNamesData As BetterArray
    Dim ColumnIndexData As BetterArray

    Dim VarNamesLLData As BetterArray
    Dim ColumnIndexLLData As BetterArray

    Dim sActSht As String
    Dim shpTemp As Worksheet
    Dim ShouldQuit As Byte
    Dim iStartSheet As Integer
    Dim iEndSheet As Integer
    Dim k As Integer 'Counter
    Dim iRow As Integer
    Dim sVarName As String 'Varname value
    Dim sVarControlType  As String 'Control of a varname
    Dim iVarIndex  As Integer   'Index of a variable
    Dim sVal As String

    Dim sPath As String
    Dim hasData As Boolean
    hasData = False


    sActSht = ActiveSheet.Name

    'Proceed by loading the files
    sPath = LoadFile("*.xlsx, *.xlsb")

    On Error GoTo errImport

    If sPath = vbNullString Then
        MsgBox TranslateLLMsg("MSG_PathImport"), vbOKOnly, TranslateLLMsg("MSG_Imports")
        Exit Sub
    End If

    hasData = LLhasData() 'Here we know if data is cleared or Not

    BeginWork xlsapp:=Application
    Application.EnableEvents = False
    Set WkbImp = Workbooks.Open(sPath)

    'Get All the sheets in the linelist
    Set TabSheetLL = GetDictionaryColumn(C_sDictHeaderSheetName)
    Set ColumnIndexData = GetDictionaryColumn(C_sDictHeaderIndex)
    Set VarNamesData = GetDictionaryColumn(C_sDictHeaderVarName)
    Set VarNamesLLData = New BetterArray
    Set ColumnIndexLLData = New BetterArray
    Set ImportVarData = New BetterArray

    'For each sheet in the imported workbook, if the sheet is in the linelist
    'Import the data the two sheets, keeping in mind We can add at the end, or
    'Just at the begining

    k = 1
    ImportReport = False

    'Set and update the temp sheet for Edition
    Set shpTemp = ThisWorkbook.Worksheets(C_sSheetImportTemp)
    sVal = shpTemp.Cells(1, 9)
    shpTemp.Cells.Clear
    shpTemp.Cells(1, 8).value = Format(Now, "yyyy-mm-dd Hh:Nn")
    If sVal <> vbNullString Then shpTemp.Cells(1, 9).value = sVal

    For Each shImp In WkbImp.Worksheets
        If TabSheetLL.Includes(shImp.Name) Then
            VarNamesLLData.Clear
            ColumnIndexLLData.Clear

            iStartSheet = TabSheetLL.IndexOf(shImp.Name)
            iEndSheet = TabSheetLL.LastIndexOf(shImp.Name) + 1

            VarNamesLLData.Items = VarNamesData.Slice(iStartSheet, iEndSheet)
            ColumnIndexLLData.Items = ColumnIndexData.Slice(iStartSheet, iEndSheet)

            Call ImportSheetData(shImp.Name, shImp, hasData, ColumnIndexLLData, VarNamesLLData)
        Else
            'Set import report to true
            If Not ImportReport Then ImportReport = True
            'Test if this is valid worksheet before writing
            If EnsureGoodSheetName(shImp.Name, TrimName:=False) = shImp.Name Then
                shpTemp.Cells(k, 1).value = shImp.Name
                k = k + 1
            End If
        End If
    Next

    'Test if there are variables in Linelist not in imported sheet
    For k = 1 To VarNamesData.Length
        sVarName = VarNamesData.Item(k)
        iVarIndex = ColumnIndexData.Item(k)

        sVarControlType = ThisWorkbook.Worksheets(TabSheetLL.Item(k)).Cells(C_eStartLinesLLMainSec - 1, iVarIndex).value

        If ImportVarData.Length > 0 And Not ImportVarData.Includes(sVarName) And sVarControlType <> C_sDictControlForm Then
            'Update report status
            If Not ImportReport Then ImportReport = True
            With shpTemp
                iRow = .Cells(.Rows.Count, 6).End(xlUp).Row
                iRow = iRow + 1
                .Cells(iRow, 6).value = VarNamesData.Item(k)
                .Cells(iRow, 7).value = TabSheetLL.Item(k)
            End With
        End If
    Next

    ThisWorkbook.Worksheets(sActSht).Activate
    WkbImp.Close savechanges:=False

    Set WkbImp = Nothing
    Set ColumnIndexData = Nothing
    Set VarNamesData = Nothing
    Set VarNamesLLData = Nothing
    Set ColumnIndexLLData = Nothing

    EndWork xlsapp:=Application
    Application.EnableEvents = True


    If Not ImportReport Then
        ShouldQuit = MsgBox(TranslateLLMsg("MSG_FinishImport"), vbQuestion + vbYesNo, TranslateLLMsg("MSG_Imports"))

        If ShouldQuit = vbYes Then
            F_ImportMig.Hide
        End If
    Else
        ShouldQuit = MsgBox(TranslateLLMsg("MSG_FinishImportRep"), vbQuestion + vbYesNo, TranslateLLMsg("MSG_Imports"))

        If ShouldQuit = vbYes Then
            F_ImportMig.Hide
            Call ShowImportReport
        End If
    End If

    Exit Sub

errImport:
    MsgBox TranslateLLMsg("MSG_ErrorImport"), vbCritical + vbOKOnly, TranslateLLMsg("MSG_Imports")
    EndWork xlsapp:=Application
    Application.EnableEvents = True
    Exit Sub

End Sub

' Show Import Report ---------------------------------------------------------

Sub ShowImportReport()

    Dim TabRep As BetterArray
    Dim shp As Worksheet
    Dim iRow As Integer
    Dim i As Integer

    Set TabRep = New BetterArray
    Set shp = ThisWorkbook.Worksheets(C_sSheetImportTemp)

    'Sheet not found
    With shp
        iRow = .Cells(Rows.Count, 1).End(xlUp).Row

        If iRow >= 1 Then
            TabRep.FromExcelRange .Range(.Cells(1, 1), .Cells(1, iRow))
            F_ImportRep.LST_ImpRepSheet.ColumnCount = 1
            F_ImportRep.LST_ImpRepSheet.List = TabRep.Items
        End If

        'Variable not imported
        iRow = .Cells(.Rows.Count, 3).End(xlUp).Row

        If iRow >= 1 Then
            TabRep.Clear
            TabRep.FromExcelRange .Range(.Cells(1, 3), .Cells(iRow, 4))
            F_ImportRep.LST_ImpRepVarImp.ColumnCount = 2
            F_ImportRep.LST_ImpRepVarImp.List = TabRep.Items
        End If

        iRow = .Cells(.Rows.Count, 6).End(xlUp).Row

        If iRow >= 1 Then
            TabRep.Clear
            TabRep.FromExcelRange .Range(.Cells(1, 6), .Cells(iRow, 7))
            F_ImportRep.LST_ImpRepVarLL.ColumnCount = 2
            F_ImportRep.LST_ImpRepVarLL.List = TabRep.Items
        End If

        If .Cells(1, 8).value <> vbNullString Then
            F_ImportRep.TXT_ImportRepData.value = TranslateLLMsg("MSG_ImportDone") & " " & .Cells(1, 8).value
        Else
             F_ImportRep.TXT_ImportRepData.value = TranslateLLMsg("MSG_NoImportDone")
        End If

        If .Cells(1, 9).value <> vbNullString Then
            F_ImportRep.TXT_ImportRepGeo.value = TranslateLLMsg("MSG_ImportGeoDone") & " " & .Cells(1, 9).value
        Else
             F_ImportRep.TXT_ImportRepGeo.value = TranslateLLMsg("MSG_NoImportGeoDone")
        End If

    End With

    'Show Import report
    F_ImportRep.Show
    Set TabRep = Nothing

End Sub



'----------------------------------- GEOBASE IMPORT   ------------------------------------------------


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
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 9).value = Format(Now, "yyyy-mm-dd Hh:Nn")

        ShouldQuit = MsgBox(TranslateLLMsg("MSG_FinishImportGeo"), vbQuestion + vbYesNo, "Import GeoData")

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

    ShouldQuit = MsgBox(TranslateLLMsg("MSG_FinishImportHistoricGeo"), vbQuestion + vbYesNo, "Import Historic")

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
        MsgBox TranslateLLMsg("MSG_ErrHistoricGeo")
        EndWork xlsapp:=Application
        Exit Sub

End Sub

'Clear the historic Data

Sub ClearHistoricGeobase()
    Dim WkshGeo As Worksheet
    Dim ShouldDelete As Integer

    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    On Error GoTo errClearHistoric

    ShouldDelete = MsgBox(TranslateLLMsg("MSG_HistoricDelete"), vbExclamation + vbYesNo, TranslateLLMsg("MSG_DeleteHistoric"))

    If ShouldDelete = vbYes Then
        If Not WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
            WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange.Delete
        End If

        If Not WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
            WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange.Delete
        End If
        MsgBox TranslateLLMsg("MSG_Done"), vbInformation,  TranslateLLMsg("MSG_DeleteHistoric"))
    End If

    'Add a message to say it is done
    Set WkshGeo = Nothing

    Exit Sub

errClearHistoric:
        MsgBox TranslateLLMsg("MSG_ErrCleanHistoric")
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
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = LLSheetData.Items(i)
            sPrevSheetName = LLSheetData.Items(i)
            ExportData.Clear
            ExportData.FromExcelRange ThisWorkbook.Worksheets(sPrevSheetName).ListObjects("o" & ClearString(sPrevSheetName, False)).Range
            ExportData.ToExcelRange .Worksheets(sPrevSheetName).Cells(1, 1)
            i = i + 1
        Wend

        'Sheets of type Admin
        i = 1
        While i <= AdmSheetData.UpperBound
            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = AdmSheetData.Items(i)
            sPrevSheetName = AdmSheetData.Items(i)
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
        MsgBox TranslateLLMsg("MSG_ErrExportData")
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
        MsgBox TranslateLLMsg("MSG_ErrExportGeo")
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
        MsgBox TranslateLLMsg("MSG_ErrExportHistoricGeo")
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
            MsgBox TranslateLLMsg("MSG_PathTooLong")
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
            MsgBox TranslateLLMsg("MSG_NoExport"), vbCritical, TranslateLLMsg("MSG_Migration")
            Exit Sub
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

        ShouldQuit = MsgBox(TranslateLLMsg("MSG_FinishedExports"), vbQuestion + vbYesNo, TranslateLLMsg("MSG_Migration"))
        If ShouldQuit = vbYes Then
            F_ExportMig.Hide
        End If

    End If

    Set ExportPath = Nothing

    Exit Sub

ErrPath:
        MsgBox TranslateLLMsg("MSG_ErrExportPath"), vbCritical + vbOKOnly, TranslateLLMsg("MSG_Migration")
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
