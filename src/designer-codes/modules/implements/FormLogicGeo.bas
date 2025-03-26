Attribute VB_Name = "FormLogicGeo"
Attribute VB_Description = "Form implementation of GeoApp"

'@ModuleDescription("Form implementation of GeoApp")
'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable, ImplicitActiveSheetReference

'Unused variables should be ignored because this module is copied to the geo form

Option Explicit

Private Const GEOSHEET As String = "Geo"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const SEP As String = " | " 'This is the separator for values in the text box at the bottom
Private Const NACHAR As String = " | N/A"
Private Const NACHARREV As String = "N/A | "

Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation
Private geo As ILLGeo
Private hfOrGeo As Byte

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set lltranssh = wb.Worksheets(LLSHEET)
    Set dicttranssh = wb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
End Sub

Private Sub InitializeElements()
    'Initialize hf or geo
    If Me.FRM_Geo.Visible Then
        hfOrGeo = GeoScopeAdmin
    Else
        hfOrGeo = GeoScopeHF
    End If
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Search values in the search box
'concatenateGeoTable is The concatenate data
'searchedvalue is the string to searrch
'1 - search in the concatenated table
'2- Add values where there are some matches in another table
'3- Render the table if it is not empty
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub searchValue(ByVal searchedValue As String, _
                Optional ByVal scope As Byte = 0, _
                Optional ByVal onHistoric As Boolean = False)

    Dim resultTable As BetterArray
    Set resultTable = New BetterArray
    Dim counter As Long
    Dim lstObj As Object 'Form List Object
    Dim concatTab As BetterArray

    Set concatTab = New BetterArray

    If scope = GeoScopeAdmin Then
        'You can search either on historic, or on aggregated values
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

    'Create a table of the found values (called resultTable)
    If Len(searchedValue) >= 3 Then

        For counter = concatTab.LowerBound To concatTab.UpperBound
            If InStr(1, LCase(concatTab.Item(counter)), LCase(searchedValue)) > 0 Then
                resultTable.Push concatTab.Item(counter)
            End If
        Next

        'Render the table if some values are found
        If resultTable.Length > 0 Then
            resultTable.Sort
            lstObj.List = resultTable.Items
        Else
            'If Not, check if there have been some input in the concat and render
            lstObj.Clear
        End If
    Else
        lstObj.List = concatTab.Items
    End If
End Sub

'This command is at the end, when you close the geoapp
'It basically update all the required data and input selected data in
'the linelist worksheet
Private Sub CMD_Copier_Click()

    Dim tempTable As BetterArray 'Temporary table for data manipulation purposes
    Dim selectedValue As String
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim hRng As Range
    Dim nbOffset As Long
    Dim calcRng As Range 'Range to calculate
    Dim sheetTag As String
    Dim cellName As String
    Dim selectedRng As Range 'Selected Range to fill the values with
    Dim nbLines As Long
    
    On Error GoTo ErrGeo
    InitializeElements
    
    selectedValue = Me.TXT_Msg.Value

    'Exit if nothing is selected
    If selectedValue = vbNullString Then
        Me.Hide
        Exit Sub
    End If

    Set cellRng = ActiveCell 'First cell for a geo value
    On Error Resume Next
        Set selectedRng = Selection
    On Error GoTo 0

    Set sh = ActiveSheet 'Linelist sheet
    sheetTag = sh.Cells(1, 3).Value

    Select Case sheetTag

    Case "HList"

        Set hRng = sh.ListObjects(1).HeaderRowRange
        nbOffset = cellRng.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)

        Set tempTable = New BetterArray
        tempTable.LowerBound = 1
        
        Select Case hfOrGeo
            'In case you selected the Geo data
        Case GeoScopeAdmin
        
            'Writing the selected data in the linelist sheet
            tempTable.Items = Split(selectedValue, SEP)

            If tempTable.Length > 0 Then
                
                'Clear the cells before filling
                Application.EnableEvents = False
                
                sh.Range(cellRng, cellRng.Offset(, 3)).ClearContents
                
                'Test if a range has been selected and fill all selected value
                If Not (selectedRng Is Nothing) Then
                        nbLines = 1
                        Do While (nbLines <= selectedRng.Rows.Count)
                            'For each rows in the selection, add the values
                            tempTable.ToExcelRange Destination:=selectedRng.Cells(nbLines, 1), TransposeValues:=True
                            nbLines = nbLines + 1
                        Loop
                Else
                    tempTable.ToExcelRange Destination:=cellRng, TransposeValues:=True
                End If

                Application.EnableEvents = True
            End If
            
            geo.UpdateHistoric selectedValue, GeoScopeAdmin
            
            'In Case we are dealing with the health facility
            '(basically the same thing with little modifications)
        Case GeoScopeHF
            
            Application.EnableEvents = False

            If Not (selectedRng Is Nothing) Then
                nbLines = 1
                Do While (nbLines <= selectedRng.Rows.Count)
                    selectedRng.Cells(nbLines, 1).Value = selectedValue
                    nbLines = nbLines + 1
                Loop
            Else
                cellRng.Value = selectedValue
            End If

            Application.EnableEvents = True
            geo.UpdateHistoric selectedValue, GeoScopeHF
        End Select

        calcRng.calculate
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
            selectedValue = Application.WorksheetFunction.Trim(Replace(selectedValue, NACHAR, vbNullString))
            selectedValue = Application.WorksheetFunction.Trim(Replace(selectedValue, NACHARREV, vbNullString))
            tempTable.Items = Split(selectedValue, SEP)

            selectedValue = tempTable.ToString(separator:=SEP, OpeningDelimiter:=vbNullString, _
                                               ClosingDelimiter:=vbNullString, QuoteStrings:=False)

            Application.EnableEvents = False
            cellRng.Value = selectedValue
            On Error Resume Next
            cellName = cellRng.Name.Name
            on Error GoTo 0
            UpdateSpatioTemporalFormulas cellName, tempTable.Length 
            Application.EnableEvents = True
        End Select

        Me.TXT_Msg.Value = vbNullString
        Me.Hide
        sh.UsedRange.Calculate
        sh.UsedRange.WrapText = TRUE
        Exit Sub
    End Select

ErrGeo:
    MsgBox tradmess.TranslatedValue("MSG_ErrWriteGeo"), vbCritical + vbOKOnly
End Sub

'Clear one historic (either hf or geo) for the geobase
Private Sub ClearOneHistoricGeobase(Optional ByVal scope As Byte = 0)
    
    Dim confirm As Boolean
    Dim lstObj As Object

    confirm = (MsgBox( _
                    tradmess.TranslatedValue("MSG_DeleteOneHistoric"), _
                    vbExclamation + vbYesNo, _
                    tradmess.TranslatedValue("MSG_DeleteHistoric")) = _
                    vbYes)

    If Not confirm Then Exit Sub

    If scope = GeoScopeAdmin Then
        Set lstObj = Me.LST_Histo
    Else
        Set lstObj = Me.LST_HistoF
    End If
    
    geo.ClearHistoric scope
    lstObj.Clear
    'It is done, inform the user
    MsgBox tradmess.TranslatedValue("MSG_Done"), _
           vbInformation, _
           tradmess.TranslatedValue("MSG_DeleteHistoric")
End Sub


Private Sub CMD_GeoClearHisto_Click()
    InitializeElements
    ClearOneHistoricGeobase hfOrGeo
End Sub

'Closing the Geoapp
Private Sub CMD_Retour_Geo_Click()
    Me.Hide
End Sub

'Those are procedures to show the following list in one item is selected.
'They rely on ShowAdmin*List functions coded in the LinelistGeo module
Private Sub LST_Adm1_Click()
    ShowAdmin2List Me.LST_Adm1.Value, GeoScopeAdmin
End Sub

Private Sub LST_Adm2_Click()
    ShowAdmin3List Me.LST_Adm2.Value, GeoScopeAdmin, sep
End Sub

Private Sub LST_Adm3_Click()
    ShowAdmin4List Me.LST_Adm3.Value, GeoScopeAdmin, sep
End Sub

Private Sub LST_Adm4_Click()
    Dim selectedValue As String

    'SEP is a constant defined above, which is the separator
    selectedValue = Me.LST_Adm1.Value & SEP & _
                    Me.LST_Adm2.Value & SEP & _
                    Me.LST_Adm3.Value & SEP & _
                    Me.LST_Adm4.Value

    Me.TXT_Msg.Value = selectedValue
End Sub

Private Sub LST_AdmF1_Click()
    ShowAdmin2List Me.LST_AdmF1.Value, GeoScopeHF
End Sub

Private Sub LST_AdmF2_Click()
    ShowAdmin3List Me.LST_AdmF2.Value, GeoScopeHF, sep
End Sub

Private Sub LST_AdmF3_Click()
    ShowAdmin4List Me.LST_AdmF3.Value, GeoScopeHF, sep
End Sub

Private Sub LST_AdmF4_Click()
    Dim selectedValue As String

    'SEP is a constant defined above, which is the separator
    selectedValue = Me.LST_AdmF4.Value & SEP & _
                    Me.LST_AdmF3.Value & SEP & _
                    Me.LST_AdmF2.Value & SEP & _
                    Me.LST_AdmF1.Value
    Me.TXT_Msg.Value = selectedValue
End Sub

'Those are trigerring event for the Histo
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

Private Sub TXT_Recherche_Change()
    InitializeElements
    'Search any value in geo data
    searchValue searchedValue:=Me.TXT_Recherche.Value, _
                scope:=hfOrGeo, onHistoric:=False
End Sub

Private Sub TXT_RechercheF_Change()
    'Search any value in hf data
    InitializeElements

    searchValue searchedValue:=Me.TXT_RechercheF.Value, _
                scope:=hfOrGeo, onHistoric:=False
End Sub

Private Sub TXT_RechercheHisto_Change()
    InitializeElements

    'Search any value in historic geo data
    searchValue searchedValue:=Me.TXT_RechercheHisto.Value, _
                scope:=hfOrGeo, onHistoric:=True

End Sub

Private Sub TXT_RechercheHistoF_Change()
    InitializeElements

    'Search any value in historic facility data
    searchValue searchedValue:=Me.TXT_RechercheHistoF.Value, _
                scope:=hfOrGeo, onHistoric:=True
End Sub

'Translate the form, resize it
Private Sub UserForm_Initialize()
    
    InitializeTrads
    InitializeElements

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 650
    Me.height = 450
End Sub
