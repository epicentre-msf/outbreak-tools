Attribute VB_Name = "DesignerMainHelpers"

'Helper functions for the designerMain
Option Explicit


'Set All the Input ranges to white
Sub SetInputRangesToWhite()

    SheetMain.Range(C_sRngPathGeo).Interior.Color = vbWhite
    SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLName).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    SheetMain.Range(C_sRngEdition).Interior.Color = vbWhite

End Sub

'Control for Linelist generation
'A Control Function to be sure that everything is fine for linelist Generation
Public Function ControlForGenerate() As Boolean

    Dim bGeo As Boolean

    ControlForGenerate = False

    'Checking coherence of the Dictionnary --------------------------------------------------------

    'Be sure the dictionary path is not empty
    If SheetMain.Range(C_sRngPathDic).value = "" Then
       SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathDic")
       SheetMain.Range(C_sRngPathDic).Interior.Color = GetColor("RedEpi")
       Exit Function
    End If

    'Now check if the file exists
    If Dir(SheetMain.Range(C_sRngPathDic).value) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathDic")
        SheetMain.Range(C_sRngPathDic).Interior.Color = GetColor("RedEpi")
        Exit Function
    End If

    'Be sure the dictionnary is not opened
    If Helpers.IsWkbOpened(Dir(SheetMain.Range(C_sRngPathDic).value)) Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseDic")
        SheetMain.Range(C_sRngPathDic).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_AlreadyOpen"), vbExclamation + vbOKOnly, TranslateMsg("MSG_Title_Dictionnary")
        Exit Function
    End If

    SheetMain.Range(C_sRngPathDic).Interior.Color = GetColor("White") 'if path is OK

    'Checking coherence of the GEO  ------------------------------------------------

    'Be sure the geo path is not empty
    If SheetMain.Range(C_sRngPathGeo).value = "" Then
       SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
       SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
       MsgBox TranslateMsg("MSG_PathGeo"), vbExclamation + vbOKOnly, TranslateMsg("MSG_TitleGeo")
       Exit Function
    End If

    'Now check if the file exists
    If Dir(SheetMain.Range(C_sRngPathGeo).value) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
         MsgBox TranslateMsg("MSG_PathGeo"), vbExclamation + vbOKOnly, TranslateMsg("MSG_TitleGeo")
        Exit Function
    End If

    bGeo = (SheetGeo.ListObjects(C_sTabadm1).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabAdm2).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabAdm3).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabAdm4).DataBodyRange Is Nothing)

    'Be sure the geo has been loaded correctly ie the geo data is not empty
    If bGeo Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LoadGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
        MsgBox TranslateMsg("MSG_GeoNotLoaded"), vbExclamation + vbOKOnly, TranslateMsg("MSG_TitleGeo")
        Exit Function
    End If

    SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("White") 'if path is OK

    'Checking coherence of the Linelist File ------------------------------------------------------

    'Be sure the linelist directory is not empty
    If SheetMain.Range(C_sRngLLDir).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_PathLL"), vbExclamation + vbOKOnly, TranslateMsg("MSG_TitleLL")
        Exit Function
    End If

    'Be sure the dictionnary is not opened
    If Helpers.IsWkbOpened(Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb")) Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseOutPut")
        SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_CloseOutPut"), vbExclamation + vbOKOnly, TranslateMsg("MSG_Title_OutPut")
        Exit Function
    End If

    'Be sure the directory for the linelist exists
    'Seems like this step is not working on Mac
    If Dir(SheetMain.Range(C_sRngLLDir).value & "*", vbDirectory) = vbNullString Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_PathLL"), vbExclamation + vbOKOnly, TranslateMsg("MSG_TitleLL")
        Exit Function
    End If

    SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("White") 'if path is OK

    'Checking coherence of the linelist name ------------------------------------------------------

    'be sure the linelist name is not empty
    If SheetMain.Range(C_sRngLLName) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLName")
        SheetMain.Range(C_sRngLLName).Interior.Color = GetColor("RedEpi")
        Exit Function
    End If

    'Be sure the linelist workbook is not already opened
    If Helpers.IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = GetColor("RedEpi")
        Exit Function
    End If

    'Be sure the linelist name is well written

    SheetMain.Range(C_sRngLLName).value = FileNameControl(SheetMain.Range(C_sRngLLName).value)

    'Call SetInputRangesToWhite
    ControlForGenerate = True

End Function


Function FileNameControl(ByVal sName As String) As String
'In the file name, replace forbidden characters with an underscore

    FileNameControl = vbNullString

    sName = Replace(sName, "<", "_")
    sName = Replace(sName, ">", "_")
    sName = Replace(sName, ":", "_")
    sName = Replace(sName, "|", "_")
    sName = Replace(sName, "?", "_")
    sName = Replace(sName, "/", "_")
    sName = Replace(sName, "\", "_")
    sName = Replace(sName, "*", "_")
    sName = Replace(sName, ".", "_")
    sName = Replace(sName, """", "_")

    FileNameControl = Application.WorksheetFunction.Trim(sName)

End Function

'Prepare temporary folder for the linelist generation process, to avoid
'conflicts with various files

Public Sub PrepareTemporaryFolder(Optional Create As Boolean = True)

'required temporary folder for analysis
    On Error Resume Next
        Workbooks("Temp.xlsb").Close SaveChanges:=False
        Workbooks("Temp").Close SaveChanges:=False
        Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "Temp.xlsb"
        Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "*.frm"
        Kill SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "*.frx"
        RmDir SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_"
        If Create Then MkDir SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" 'create a folder for sending all the data from designer
    On Error GoTo 0

End Sub

Public Sub AddTableNames()
    Dim iCol As Long
    Dim iRow As Long
    Dim iSheetNameCol As Integer
    Dim sPrevSheetName As String
    Dim sTableName As String
    Dim DictHeaders As BetterArray
    Dim i As Long
    Dim iTableIndex As Long

    Set DictHeaders = New BetterArray
    Set DictHeaders = GetHeaders(ThisWorkbook, C_sParamSheetDict, 1)

    iSheetNameCol = DictHeaders.IndexOf(C_sDictHeaderSheetName)

    With ThisWorkbook.Worksheets(C_sParamSheetDict)
        iRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        iCol = DictHeaders.Length + 1
        iTableIndex = 1

        sPrevSheetName = .Cells(2, iSheetNameCol).value
        sTableName = "table" & iTableIndex
        Set DictHeaders = Nothing

        .Cells(1, iCol).value = C_sDictHeaderTableName
        For i = 2 To iRow
            If sPrevSheetName <> .Cells(i, iSheetNameCol).value Then
                'New sheet name, new table
                sPrevSheetName = .Cells(i, iSheetNameCol).value
                iTableIndex = iTableIndex + 1
                sTableName = "table" & iTableIndex
            End If

            .Cells(i, iCol).value = sTableName
        Next
    End With
End Sub



'Move analysis Data from the analysis Sheet to the DesignerWorkbook
    Public Function MoveAnalysis(SrcWkb As Workbook)

        Dim DesRng As Range 'Range to resize the new list object in the designer
        Dim SetupRng As Range 'Range in the setup file

        Dim iPasteRow As Long
        Dim iPasteColumn As Long
        Dim iLastRow As Long
        Dim iLastColumn As Long

        Dim SetupWksh As Worksheet
        Dim DesWksh As Worksheet

        Dim Lo As ListObject

        If Not SheetExistsInWkb(SrcWkb, C_sSheetAnalysis) Then Exit Function

        Set SetupWksh = SrcWkb.Worksheets(C_sSheetAnalysis)
        Set DesWksh = DesignerWorkbook.Worksheets(C_sSheetAnalysis)

        DesWksh.Cells.Clear

        For Each Lo In SetupWksh.ListObjects

            iPasteRow = Lo.Range.Row
            iPasteColumn = Lo.Range.Column

            SetupWksh.Cells(iPasteRow - 2, iPasteColumn).Copy DesWksh.Cells(iPasteRow - 2, iPasteColumn)

            'Find where data is entered from the first column
            iLastRow = SetupWksh.Cells(iPasteRow, iPasteColumn).End(xlDown).Row
            iLastColumn = SetupWksh.Cells(iPasteRow, iPasteColumn).End(xlToRight).Column

            With SetupWksh
                Set SetupRng = .Range(.Cells(iPasteRow, iPasteColumn), .Cells(iLastRow, iLastColumn))
            End With

            With DesWksh
                Set DesRng = .Range(.Cells(iPasteRow, iPasteColumn), .Cells(iLastRow, iLastColumn))
                DesRng.value = SetupRng.value
                .ListObjects.Add(xlSrcRange, DesRng, , xlYes).Name = Lo.Name
            End With

        Next


        Set DesRng = Nothing
        Set SetupRng = Nothing
        Set Lo = Nothing
        Set SetupWksh = Nothing
        Set DesWksh = Nothing
    End Function


'update the progress status
Sub StatusBar_Updater(sCpte As Single)

        Dim CurrentStatus As Integer
        Dim pctDone As Integer
        Dim bCurrEvent As Boolean

        bCurrEvent = Application.ScreenUpdating

        Application.ScreenUpdating = True

        CurrentStatus = (C_iNumberOfBars) * Round(sCpte / 100, 1)
        SheetMain.Range(C_sRngUpdate).value = "[" & String(CurrentStatus, "|") & Space(C_iNumberOfBars - CurrentStatus) & "]" & " " & CInt(sCpte) & "% " & TranslateMsg("MSG_BuildLL")

        Application.ScreenUpdating = bCurrEvent

End Sub


