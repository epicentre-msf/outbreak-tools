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

    SheetMain.Range(C_sRngLLName).Interior.Color = GetColor("White") 'If path is OK
    ControlForGenerate = True

End Function
