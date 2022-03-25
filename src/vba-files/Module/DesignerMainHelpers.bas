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


'Cancel Generation and stop all processes.

Sub CancelGenerate()
    Dim answer As Integer
    
    answer = MsgBox(TranslateMsg("MSG_ConfCancel"), vbYesNo)
    
    
    If answer = vbYes Then
        Call SetInputRangesToWhite
        
        ShowHideCmdValidation show:=False
        SheetMain.Shapes("SHP_OpenLL").Visible = msoFalse
        End 'This is probably to avoid, but will come back later on that
    End If
    
    MsgBox TranslateMsg("MSG_Continue")
End Sub

'Show/Hide the shapes for linelist creation
Public Sub ShowHideCmdValidation(show As Boolean)

    SheetMain.Shapes("SHP_Generer").Visible = show
    SheetMain.Shapes("SHP_Annuler").Visible = show
    SheetMain.Shapes("SHP_CtrlNouv").Visible = Not show
End Sub


'Control for Linelist generation
'A Control Function to be sure that everything is fine for linelist Generation
Public Function ControlForGenerate() As Boolean
    
    Dim bGeo As Boolean

    ControlForGenerate = False
    'Hide the shapes for linelist generation
    ShowHideCmdValidation show:=False
    
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
        MsgBox TranslateMsg("MSG_AlreadyOpen"), vbExclamation + vbokOnly, TranslateMsg("MSG_Title_Dictionnary")
        Exit Function
    End If
    
    SheetMain.Range(C_sRngPathDic).Interior.Color = GetColor("White") 'if path is OK
    
    'Checking coherence of the GEO  ------------------------------------------------
    
    'Be sure the geo path is not empty
    If SheetMain.Range(C_sRngPathGeo).value = "" Then
       SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
       SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
       MsgBox TranslateMsg("MSG_PathGeo"), vbExclamation + vbokOnly, TranslateMsg("MSG_TitleGeo")
       Exit Function
    End If
    
    'Now check if the file exists
    If Dir(SheetMain.Range(C_sRngPathGeo).value) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
         MsgBox TranslateMsg("MSG_PathGeo"), vbExclamation + vbokOnly, TranslateMsg("MSG_TitleGeo")
        Exit Function
    End If

    bGeo = (SheetGeo.ListObjects(C_sTabADM1).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabADM2).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabADM3).DataBodyRange Is Nothing) Or _
            (SheetGeo.ListObjects(C_sTabADM4).DataBodyRange Is Nothing)

    'Be sure the geo has been loaded correctly ie the geo data is not empty
    If bGeo Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LoadGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("RedEpi")
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
        MsgBox TranslateMsg("MSG_GeoNotLoaded"), vbExclamation + vbokOnly, TranslateMsg("MSG_TitleGeo")
        Exit Function
    End If

    SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("White") 'if path is OK
    
    'Checking coherence of the Linelist File ------------------------------------------------------
    
    'Be sure the linelist directory is not empty
    If SheetMain.Range(C_sRngLLDir).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_PathLL"), vbExclamation + vbokOnly, TranslateMsg("MSG_TitleLL")
        Exit Function
    End If

    'Be sure the directory for the linelist exists
    If Dir(SheetMain.Range(C_sRngLLDir).value, vbDirectory) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = GetColor("RedEpi")
        MsgBox TranslateMsg("MSG_PathLL"), vbExclamation + vbokOnly, TranslateMsg("MSG_TitleLL")
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
