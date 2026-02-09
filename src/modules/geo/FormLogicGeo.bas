Attribute VB_Name = "FormLogicGeo"
Attribute VB_Description = "Form callbacks for F_Geo — delegates to GeoModule"

'@Folder("Geo")
'@ModuleDescription("Form callbacks for F_Geo -- delegates to GeoModule")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable, ImplicitActiveSheetReference, UseMeaningfulName, HungarianNotation

Option Explicit

Private Const GEOSHEET As String = "Geo"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const SEP As String = " | "
Private Const NACHAR As String = " | N/A"
Private Const NACHARREV As String = "N/A | "

Private tradform As ITranslationObject
Private tradmess As ITranslationObject
Private geo As ILLGeo
Private hfOrGeo As Byte

'@section Initialization
'===============================================================================

' @description Initialize translation objects and the LLGeo instance.
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set lltrads = LLTranslations.Create( _
        wb.Worksheets(LLSHEET), _
        wb.Worksheets(TRADSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
End Sub

' @description Determine the current scope (geo vs hf) from form visibility.
Private Sub InitializeElements()
    If Me.FRM_Geo.Visible Then
        hfOrGeo = GeoScopeAdmin
    Else
        hfOrGeo = GeoScopeHF
    End If
End Sub

'@section Search
'===============================================================================

' @description Search values in concat/historic lists and filter the form list.
Private Sub SearchValue(ByVal searchedValue As String, _
                        Optional ByVal scope As Byte = 0, _
                        Optional ByVal onHistoric As Boolean = False)

    Dim resultTable As BetterArray
    Dim counter As Long
    Dim lstObj As Object
    Dim concatTab As BetterArray

    Set resultTable = New BetterArray
    Set concatTab = New BetterArray

    If scope = GeoScopeAdmin Then
        If onHistoric Then
            Set lstObj = Me.LST_Histo
            concatTab.FromExcelRange Range("histo_geo")
        Else
            Set lstObj = Me.LST_ListeAgre
            concatTab.FromExcelRange Range("adm4_concat")
        End If
    Else
        If onHistoric Then
            Set lstObj = Me.LST_HistoF
            concatTab.FromExcelRange Range("histo_hf")
        Else
            Set lstObj = Me.LST_ListeAgreF
            concatTab.FromExcelRange Range("hf_concat")
        End If
    End If

    If Len(searchedValue) >= 3 Then
        For counter = concatTab.LowerBound To concatTab.UpperBound
            If InStr(1, LCase(concatTab.Item(counter)), LCase(searchedValue)) > 0 Then
                resultTable.Push concatTab.Item(counter)
            End If
        Next

        If resultTable.Length > 0 Then
            resultTable.Sort
            lstObj.List = resultTable.Items
        Else
            lstObj.Clear
        End If
    Else
        lstObj.List = concatTab.Items
    End If
End Sub

'@section CMD_Copier — Write Selected Value
'===============================================================================

' @description Write the selected geo/hf value to the active cell in the linelist.
Private Sub CMD_Copier_Click()
    Dim tempTable As BetterArray
    Dim selectedValue As String
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim hRng As Range
    Dim nbOffset As Long
    Dim calcRng As Range
    Dim sheetTag As String
    Dim cellName As String
    Dim selectedRng As Range
    Dim nbLines As Long

    On Error GoTo ErrGeo
    InitializeElements

    selectedValue = Me.TXT_Msg.Value
    If selectedValue = vbNullString Then
        Me.Hide
        Exit Sub
    End If

    Set cellRng = ActiveCell
    On Error Resume Next
    Set selectedRng = Selection
    On Error GoTo ErrGeo

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    Select Case sheetTag

    Case "HList"
        Set hRng = sh.ListObjects(1).HeaderRowRange
        nbOffset = cellRng.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)

        Set tempTable = New BetterArray
        tempTable.LowerBound = 1

        Select Case hfOrGeo
        Case GeoScopeAdmin
            tempTable.Items = Split(selectedValue, SEP)

            If tempTable.Length > 0 Then
                Application.EnableEvents = False
                sh.Range(cellRng, cellRng.Offset(, 3)).ClearContents

                If Not (selectedRng Is Nothing) Then
                    nbLines = 1
                    Do While nbLines <= selectedRng.Rows.Count
                        tempTable.ToExcelRange Destination:=selectedRng.Cells(nbLines, 1), _
                                               TransposeValues:=True
                        nbLines = nbLines + 1
                    Loop
                Else
                    tempTable.ToExcelRange Destination:=cellRng, TransposeValues:=True
                End If

                Application.EnableEvents = True
            End If

            geo.UpdateHistoric selectedValue, GeoScopeAdmin

        Case GeoScopeHF
            Application.EnableEvents = False

            If Not (selectedRng Is Nothing) Then
                nbLines = 1
                Do While nbLines <= selectedRng.Rows.Count
                    selectedRng.Cells(nbLines, 1).Value = selectedValue
                    nbLines = nbLines + 1
                Loop
            Else
                cellRng.Value = selectedValue
            End If

            Application.EnableEvents = True
            geo.UpdateHistoric selectedValue, GeoScopeHF
        End Select

        calcRng.Calculate
        Me.TXT_Msg.Value = vbNullString
        Me.Hide
        Exit Sub

    Case "SPT-Analysis"
        Select Case hfOrGeo
        Case GeoScopeHF
            Application.EnableEvents = False
            cellRng.Value = selectedValue
            Application.EnableEvents = True

        Case GeoScopeAdmin
            Set tempTable = New BetterArray
            selectedValue = Application.WorksheetFunction.Trim( _
                Replace(selectedValue, NACHAR, vbNullString))
            selectedValue = Application.WorksheetFunction.Trim( _
                Replace(selectedValue, NACHARREV, vbNullString))
            tempTable.Items = Split(selectedValue, SEP)

            selectedValue = tempTable.ToString(separator:=SEP, _
                OpeningDelimiter:=vbNullString, _
                ClosingDelimiter:=vbNullString, QuoteStrings:=False)

            Application.EnableEvents = False
            cellRng.Value = selectedValue
            On Error Resume Next
            cellName = cellRng.Name.Name
            On Error GoTo ErrGeo
            UpdateSpatioTemporalFormulas cellName, tempTable.Length
            Application.EnableEvents = True
        End Select

        Me.TXT_Msg.Value = vbNullString
        Me.Hide
        sh.UsedRange.Calculate
        sh.UsedRange.WrapText = True
        Exit Sub
    End Select

ErrGeo:
    MsgBox tradmess.TranslatedValue("MSG_ErrWriteGeo"), vbCritical + vbOKOnly
End Sub

'@section Historic
'===============================================================================

' @description Clear one historic list (geo or hf) with user confirmation.
Private Sub ClearOneHistoricGeobase(Optional ByVal scope As Byte = 0)
    Dim confirm As Boolean
    Dim lstObj As Object

    confirm = (MsgBox( _
        tradmess.TranslatedValue("MSG_DeleteOneHistoric"), _
        vbExclamation + vbYesNo, _
        tradmess.TranslatedValue("MSG_DeleteHistoric")) = vbYes)

    If Not confirm Then Exit Sub

    If scope = GeoScopeAdmin Then
        Set lstObj = Me.LST_Histo
    Else
        Set lstObj = Me.LST_HistoF
    End If

    geo.ClearHistoric scope
    lstObj.Clear

    MsgBox tradmess.TranslatedValue("MSG_Done"), _
           vbInformation, _
           tradmess.TranslatedValue("MSG_DeleteHistoric")
End Sub

Private Sub CMD_GeoClearHisto_Click()
    InitializeElements
    ClearOneHistoricGeobase hfOrGeo
End Sub

'@section Navigation
'===============================================================================

Private Sub CMD_Retour_Geo_Click()
    Me.Hide
End Sub

'@section Admin List Click Events — Delegates to GeoModule
'===============================================================================

Private Sub LST_Adm1_Click()
    ShowAdmin2List Me.LST_Adm1.Value, GeoScopeAdmin
End Sub

Private Sub LST_Adm2_Click()
    ShowAdmin3List Me.LST_Adm2.Value, GeoScopeAdmin, SEP
End Sub

Private Sub LST_Adm3_Click()
    ShowAdmin4List Me.LST_Adm3.Value, GeoScopeAdmin, SEP
End Sub

Private Sub LST_Adm4_Click()
    Me.TXT_Msg.Value = Me.LST_Adm1.Value & SEP & _
                        Me.LST_Adm2.Value & SEP & _
                        Me.LST_Adm3.Value & SEP & _
                        Me.LST_Adm4.Value
End Sub

Private Sub LST_AdmF1_Click()
    ShowAdmin2List Me.LST_AdmF1.Value, GeoScopeHF
End Sub

Private Sub LST_AdmF2_Click()
    ShowAdmin3List Me.LST_AdmF2.Value, GeoScopeHF, SEP
End Sub

Private Sub LST_AdmF3_Click()
    ShowAdmin4List Me.LST_AdmF3.Value, GeoScopeHF, SEP
End Sub

Private Sub LST_AdmF4_Click()
    Me.TXT_Msg.Value = Me.LST_AdmF4.Value & SEP & _
                        Me.LST_AdmF3.Value & SEP & _
                        Me.LST_AdmF2.Value & SEP & _
                        Me.LST_AdmF1.Value
End Sub

'@section Historic / Aggregate List Click Events
'===============================================================================

Private Sub LST_Histo_Click()
    Me.TXT_Msg.Value = Me.LST_Histo.Value
End Sub

Private Sub LST_HistoF_Click()
    Me.TXT_Msg.Value = Me.LST_HistoF.Value
End Sub

Private Sub LST_ListeAgre_Click()
    Me.TXT_Msg.Value = Me.LST_ListeAgre.Value
End Sub

Private Sub LST_ListeAgreF_Click()
    Me.TXT_Msg.Value = Me.LST_ListeAgreF.Value
End Sub

'@section Search Text Change Events
'===============================================================================

Private Sub TXT_Recherche_Change()
    InitializeElements
    SearchValue searchedValue:=Me.TXT_Recherche.Value, _
                scope:=hfOrGeo, onHistoric:=False
End Sub

Private Sub TXT_RechercheF_Change()
    InitializeElements
    SearchValue searchedValue:=Me.TXT_RechercheF.Value, _
                scope:=hfOrGeo, onHistoric:=False
End Sub

Private Sub TXT_RechercheHisto_Change()
    InitializeElements
    SearchValue searchedValue:=Me.TXT_RechercheHisto.Value, _
                scope:=hfOrGeo, onHistoric:=True
End Sub

Private Sub TXT_RechercheHistoF_Change()
    InitializeElements
    SearchValue searchedValue:=Me.TXT_RechercheHistoF.Value, _
                scope:=hfOrGeo, onHistoric:=True
End Sub

'@section UserForm Initialization
'===============================================================================

' @description Translate the form and set initial dimensions.
Private Sub UserForm_Initialize()
    InitializeTrads
    InitializeElements

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 650
    Me.Height = 450
End Sub
