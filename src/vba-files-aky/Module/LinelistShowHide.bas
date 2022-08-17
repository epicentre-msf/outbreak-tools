Attribute VB_Name = "LinelistShowHide"

Option Explicit
' Retrieving the heading of the dictionnary (names of columns)
Public TriggerShowHide As Boolean

'Show/Hidden Variables
Dim sShown As String
Dim sHidden As String
Dim sShowMandatory As String
Dim sMandatory As String
Dim sDesHidden As String


'Initialize the Hidden/Mandatory variables
Sub SetHiddenMandatory()

    sShown = TranslateLLMsg("MSG_Show")  ' translate_LineList("Show", Sheets("linelist-translation").[T_tradRange])
    sHidden = TranslateLLMsg("MSG_Hide") 'translate_LineList("Hidden", Sheets("linelist-translation").[T_tradRange])
    sMandatory = TranslateLLMsg("MSG_Mandatory") 'translate_LineList("mandatory", Sheets("linelist-translation").[T_tradRange])
    sShowMandatory = TranslateLLMsg("MSG_ShowMandatory") 'translate_LineList("Show/Mandatory", Sheets("linelist-translation").[T_tradRange])
    sDesHidden = TranslateLLMsg("MSG_DesHidden")

End Sub


Function CreateDicTitle() As BetterArray
    Dim T_DictHeaders As BetterArray                 'headers: colnames of the dictionary
    Set T_DictHeaders = New BetterArray

    'loading headers
    T_DictHeaders.Clear
    T_DictHeaders.FromExcelRange Sheets(C_sParamSheetDict).Range("A1"), DetectLastRow:=False, DetectLastColumn:=True
    'Checking the visibility variable
    If Not T_DictHeaders.Includes(C_sDictHeaderVisibility) Then
        T_DictHeaders.Push C_sDictHeaderVisibility
        'add the visibility
        ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(1, T_DictHeaders.UpperBound).value = C_sDictHeaderVisibility
    End If

    Set CreateDicTitle = T_DictHeaders.Clone
End Function

'This command loads variables and
'put all of them in the list of the show/hide forms
'only not hidden variables are shown. We need to filtered out
'those variables

Sub ClicCmdShowHide()

    Dim T_mainlab As BetterArray                 'main label table
    Dim T_varname As BetterArray                 'varname table
    Dim T_status As BetterArray                  'status table
    Dim T_DictHeaders As BetterArray                 'headers of the dictionary table
    Dim T_data As BetterArray                    'temporary data for storing the values
    Dim Wksh As Worksheet                         'Setting a temporary variable for dictionary selection
    Dim i As Integer
    Dim bremoveFromGeo As Boolean

    'Setting and initializing the tables
    Set T_mainlab = New BetterArray
    Set T_varname = New BetterArray
    Set T_status = New BetterArray
    Set T_data = New BetterArray
    Set T_DictHeaders = New BetterArray

    T_varname.LowerBound = 1
    T_mainlab.LowerBound = 1
    T_status.LowerBound = 1
    T_DictHeaders.LowerBound = 1
    T_data.LowerBound = 1

    ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

    On Error GoTo errShowHide

    Set Wksh = ThisWorkbook.Worksheets(C_sParamSheetDict)

    'Set Hidden and Show values depending on the language
    Call SetHiddenMandatory

    'Get the headers
    Set T_DictHeaders = CreateDicTitle
    'Now update the mainlabel, status and variable name

    i = 1
    bremoveFromGeo = False

     While (i <= Wksh.Cells(Wksh.Rows.Count, 1).End(xlUp).Row)

        If ActiveSheet.Name = Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderSheetName)) Then
            bremoveFromGeo = Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "2" Or _
                             Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "3" Or _
                             Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderControl)) = C_sDictControlGeo & "4"

            'update only on non hidden variables
            If Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderStatus)).value <> C_sDictStatusHid Then

                'avoid adding the other Geos
                If Not bremoveFromGeo Then
                    T_mainlab.Push Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderMainLab)).value
                    T_varname.Push Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderVarName)).value

                    If LCase(Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderStatus)).value) = C_sDictStatusMan Or _
                        Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderIndex)).value <= C_iLLSplitColumn Then
                        T_status.Push sMandatory
                        Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderVisibility)).value = sMandatory
                    ElseIf Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderVisibility)).value = sHidden Then
                        T_status.Push sHidden
                    Else
                        T_status.Push sShown
                        Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderVisibility)).value = sShown
                    End If
                End If

            Else
                Wksh.Cells(i, T_DictHeaders.IndexOf(C_sDictHeaderVisibility)).value = sDesHidden
            End If
        End If
        i = i + 1
    Wend

    T_data.Item(1) = T_mainlab.Items
    T_data.Item(2) = T_varname.Items
    T_data.Item(3) = T_status.Items



    Application.EnableEvents = False

    T_data.ArrayType = BA_MULTIDIMENSION
    Set T_data = T_data.Clone

    T_data.Transpose
    F_NomVisible.LST_NomChamp.ColumnCount = 3
    F_NomVisible.LST_NomChamp.BoundColumn = 2
    F_NomVisible.LST_NomChamp.List = T_data.Items
    'Setting objects to nothing



    Application.EnableEvents = True

    F_NomVisible.FRM_AffMas.Visible = True
    F_NomVisible.FRM_AffMas.Width = 90
    F_NomVisible.Width = 450
    F_NomVisible.Height = 270
    F_NomVisible.CMD_Fermer.SetFocus
    F_NomVisible.Show

    Call ProtectSheet

    Exit Sub

errShowHide:
        MsgBox TranslateLLMsg("MSG_UnableShowHide")
        Call ProtectSheet
        EndWork xlsapp:=Application
        Exit Sub
End Sub


'This sub will works with the logic related to the selection of oneline in the
'Show/hide multibox page
Sub UpdateVisibilityStatus(iIndex As Integer)

    Dim T_FormData As BetterArray                'Actual form data
    Set T_FormData = New BetterArray
    T_FormData.LowerBound = 1
    T_FormData.Items = F_NomVisible.LST_NomChamp.List

    BeginWork xlsapp:=Application
    Application.EnableEvents = False


    F_NomVisible.FRM_AffMas.Visible = True
    Select Case T_FormData.Items(iIndex + 1, 3)
    Case sMandatory
        TriggerShowHide = False
        F_NomVisible.OPT_Affiche.value = 1
        F_NomVisible.OPT_Affiche.Caption = sShowMandatory
        F_NomVisible.OPT_Affiche.Width = 80
        F_NomVisible.OPT_Affiche.Left = 0
        F_NomVisible.OPT_Affiche.Top = 20

        F_NomVisible.OPT_Masque.Visible = False
    Case sHidden                                'It is hidden, show masking
        TriggerShowHide = False
        F_NomVisible.OPT_Affiche.value = 0
        F_NomVisible.OPT_Affiche.Caption = sShown
        F_NomVisible.OPT_Affiche.Width = 70
        F_NomVisible.OPT_Affiche.Left = 10
        F_NomVisible.OPT_Affiche.Top = 6

        F_NomVisible.OPT_Masque.Visible = True
        F_NomVisible.OPT_Affiche.Visible = True
        F_NomVisible.OPT_Masque.value = 1
    Case Else
        TriggerShowHide = False                                   'It is shown if not
        F_NomVisible.OPT_Affiche.value = 1
        F_NomVisible.OPT_Affiche.Caption = sShown
        F_NomVisible.OPT_Affiche.Width = 70
        F_NomVisible.OPT_Affiche.Left = 10
        F_NomVisible.OPT_Affiche.Top = 6

        F_NomVisible.OPT_Masque.Visible = True
        F_NomVisible.OPT_Affiche.Visible = True
        F_NomVisible.OPT_Masque.value = 0
    End Select

    TriggerShowHide = True
    Application.EnableEvents = True
    EndWork xlsapp:=Application
End Sub

'This procedures hides or shows one column from the One sheet given the variable name selected
'in the visibility form
Sub ShowHideColumnSheet(sSheetName As String, ByVal sVarName As String, Optional bhide As Boolean = True)
    'bhide is a boolean to hide or show one column
    Dim indexCol As Integer                      'Column The index of the column to Hide
    Dim T_DictHeaders As BetterArray                 'Temporary data for headers
    Dim sControl As String                 'Extracting the control label to be sure we can hide all the geos

    BeginWork xlsapp:=Application
    ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

    'First, Get the values of the headers names
    Set T_DictHeaders = New BetterArray
    T_DictHeaders.LowerBound = 1
    Set T_DictHeaders = GetDictDataFromCondition(C_sDictHeaderSheetName, sSheetName, True)

    indexCol = T_DictHeaders.IndexOf(sVarName)

    'Extract the control column
    sControl = GetDictColumnValue(sVarName, C_sDictHeaderControl)

    'Hidding
    If indexCol > 0 Then
        'Now hiding
        ThisWorkbook.Worksheets(sSheetName).Columns(indexCol).Hidden = bhide
        'Testing if it is a geo column and hide the followings
        If sControl = C_sDictControlGeo Then
            ThisWorkbook.Worksheets(sSheetName).Columns(indexCol + 1).Hidden = bhide
            ThisWorkbook.Worksheets(sSheetName).Columns(indexCol + 2).Hidden = bhide
            ThisWorkbook.Worksheets(sSheetName).Columns(indexCol + 3).Hidden = bhide
        End If
    End If

    Call ProtectSheet
    EndWork xlsapp:=Application

End Sub

'A simple Procedure to update the third column of the next formname to either hidden or Shown before
'moving to the logic

Sub UpdateFormData(ByRef T_table As BetterArray, index As Integer, Optional bhide As Boolean = True)
    'We can't mutate easily on a 2D array. We need first to work on updating the data from a
    '1D array. That is the purpose of the T_values BetterArray
    Dim T_values As BetterArray
    Set T_values = New BetterArray

    'Get the current row
    T_values.Items = T_table.Item(index + 1)     'the index starts at -1 on a listbox
    'Update the visibility status
    If bhide Then
        T_values.Item(3) = sHidden
    Else
        T_values.Item(3) = sShown
    End If
    'Mutate in the form table
    T_table.Item(index + 1) = T_values.Items
End Sub

'Logic behind the show/hide click
Sub ShowHideLogic(iIndex As Integer)

    On Error GoTo errShowHide

    If Not TriggerShowHide Or iIndex < 0 Then    'when the form is shown at the begining, nothing is selected and index can be -1
        Exit Sub
    Else
        BeginWork xlsapp:=Application
        Application.EnableEvents = False

        Dim T_FormData As BetterArray
        Set T_FormData = New BetterArray

        T_FormData.Items = F_NomVisible.LST_NomChamp.List
        T_FormData.LowerBound = 1

        'Update data in form
        If LCase(T_FormData.Items(iIndex + 1, 2)) <> sMandatory Then

            'For mutating, we can only use the item method. Items with s, read only values,
            'we can't set values with items

            If F_NomVisible.OPT_Masque.value Then
                '// --- Here I update the Data to show "Hidden"
                UpdateFormData T_table:=T_FormData, index:=iIndex, bhide:=True
                '//--- Actually hide the column
                ShowHideColumnSheet sSheetName:=ActiveSheet.Name, sVarName:=T_FormData.Items(iIndex + 1, 2), bhide:=True
                WriteShowHide sSheetName:=ActiveSheet.Name, sVarName:=T_FormData.Items(iIndex + 1, 2), visibility:=0
            Else
                '// --- Here I udpate the data to show "Shown"
                UpdateFormData T_table:=T_FormData, index:=iIndex, bhide:=False
                ShowHideColumnSheet sSheetName:=ActiveSheet.Name, sVarName:=T_FormData.Items(iIndex + 1, 2), bhide:=False
                WriteShowHide sSheetName:=ActiveSheet.Name, sVarName:=T_FormData.Items(iIndex + 1, 2), visibility:=1
            End If
        End If

        'Reload it back
        F_NomVisible.LST_NomChamp.Clear
        F_NomVisible.LST_NomChamp.List = T_FormData.Items
        F_NomVisible.LST_NomChamp.Selected(iIndex) = True

        Application.EnableEvents = True
        EndWork xlsapp:=Application
    End If

    Exit Sub

errShowHide:
        MsgBox TranslateLLMsg("MSG_ErrShowHide")
        Call ProtectSheet
        Application.EnableEvents = True
        EndWork xlsapp:=Application
        Exit Sub
End Sub

Sub WriteShowHide(sSheetName As String, ByVal sVarName As String, visibility As Byte)
    Dim T_DictVarnames As BetterArray
    Dim T_DictSheetNames As BetterArray
    Dim iVarnameIndex As Integer
    Dim iVisIndex As Integer

    Set T_DictVarnames = GetDictionaryColumn(C_sDictHeaderVarName)
    Set T_DictSheetNames = GetDictionaryColumn(C_sDictHeaderSheetName)
    iVisIndex = GetDictionaryIndex(C_sDictHeaderVisibility)

    If T_DictSheetNames.Includes(sSheetName) Then
        If T_DictVarnames.Includes(sVarName) Then
            T_DictVarnames.LowerBound = 2
            iVarnameIndex = T_DictVarnames.IndexOf(sVarName)
            If visibility = 0 Then
                ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(iVarnameIndex, iVisIndex).value = sHidden
            ElseIf visibility = 1 Then
                ThisWorkbook.Worksheets(C_sParamSheetDict).Cells(iVarnameIndex, iVisIndex).value = sShown
            End If
        End If
    End If


End Sub
