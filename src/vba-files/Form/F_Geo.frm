VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Geo 
   Caption         =   "GEO Apps"
   ClientHeight    =   9576.001
   ClientLeft      =   60
   ClientTop       =   -264
   ClientWidth     =   12240
   OleObjectBlob   =   "F_Geo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Geo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




























































































































































































































































Option Explicit

'This command is at the end, when you close the geoapp
'It basically update all the required data and input selected data in the linelist worksheet
Private Sub CMD_Copier_Click()

    Dim T_temp As BetterArray
    Dim selectedValue As String
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim hRng As Range
    Dim nbOffset As Long
    Dim calcRng As Range 'Range to calculate
    
    Set T_temp = New BetterArray
    T_temp.LowerBound = 1

    On Error GoTo ErrGeo
    
    selectedValue = [TXT_Msg].Value
    Set cellRng = ActiveCell 'First cell for a geo value
    Set sh = ActiveSheet 'Linelist sheet
    Set hRng = sh.ListObjects(1).HeaderRowRange
    nbOffset = cellRng.Row - hRng.Row
    Set calcRng = hRng.Offset(nbOffset)
    
    Select Case iGeoType
        'In case you selected the Geo data
    Case 0
    
        'Writing the selected data in the linelist sheet
        T_temp.Clear
        T_temp.Items = Split(selectedValue, " | ")
        If T_temp.Length > 0 Then
            'Clear the cells before filling
            Application.EnableEvents = False
            sh.Range(cellRng, cellRng.Offset(, 3)).Value = ""
            If T_temp.Length = 4 Then T_temp.Reverse
            T_temp.ToExcelRange Destination:=Range(ActiveCell.Address), TransposeValues:=True
            Application.EnableEvents = True
        End If
        
        calcRng.Calculate
        [F_Geo].TXT_Msg.Value = ""
        [F_Geo].Hide
        
        'Protecting the worksheet
                
        'updating the histo data if needed
        T_temp.Clear
        Set sh = ThisWorkbook.Worksheets("Geo")
        Set Lo = sh.ListObjects("T_HISTOGEO")
        Set LoRng = Lo.Range
        
        'only update if you don't find actual value then update
        If Not T_HistoGeo.Includes(ReverseString(selectedValue)) Then T_HistoGeo.Push ReverseString(selectedValue)
        
        'Now rewrite the histo data in the list object
        If T_HistoGeo.Length > (Lo.Range.Rows.Count - 1) Then
            T_HistoGeo.ToExcelRange Destination:=LoRng.Cells(2, 1)
            'resize the list object
            Lo.Resize sh.Range(LoRng.Cells(1, 1), LoRng.Cells(T_HistoGeo.Length + 1, 1))
            Set LoRng = Lo.DataBodyRange
            LoRng.RemoveDuplicates Columns:=1, Header:=xlYes
            LoRng.Sort key1:=LoRng, Header:=xlYes
        End If
        
        'In Case we are dealing with the health facility (basically the same thing with little modifications)
    Case 1
        Application.EnableEvents = False
        cellRng.Value = selectedValue
        Application.EnableEvents = True
        'Hide the form
        calcRng.Calculate
        [F_Geo].TXT_Msg.Value = ""
        [F_Geo].Hide
        
        'Update the listObject of historic data on health facility
        Set sh = ThisWorkbook.Worksheets("Geo")
        Set Lo = sh.ListObjects("T_HISTOHF")
        Set LoRng = Lo.Range
         
        If Not T_HistoHF.Includes(selectedValue) Then T_HistoHF.Push selectedValue
            
        'Now rewrite the histo data in the list object
        If T_HistoHF.Length > (Lo.Range.Rows.Count - 1) Then
            T_HistoHF.ToExcelRange Destination:=LoRng.Cells(2, 1)
            'resize the list object
            Lo.Resize sh.Range(LoRng.Cells(1, 1), LoRng.Cells(T_HistoHF.Length + 1, 1))
            Set LoRng = Lo.DataBodyRange
            LoRng.RemoveDuplicates Columns:=1, Header:=xlYes
            LoRng.Sort key1:=LoRng, Header:=xlYes
        End If
    End Select

    Exit Sub

ErrGeo:
    MsgBox TranslateLLMsg("MSG_ErrWriteGeo"), vbCritical + vbOKOnly
End Sub

Private Sub CMD_GeoClearHisto_Click()
    Call ClearOneHistoricGeobase(iGeoType)
End Sub

'Closing the Geoapp
Private Sub CMD_Retour_Geo_Click()
    Me.Hide
End Sub

'Those are procedures to show the following list in one item is selected.
'They rely on ShowLst* functions coded in the Geo module
Private Sub LST_Adm1_Click()
    Call ShowLst2(LST_Adm1.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_Adm2_Click()
    Call ShowLst3(LST_Adm2.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_Adm3_Click()
    Call ShowLst4(LST_Adm3.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_Adm4_Click()
    sPlaceSelection = ReverseString([F_Geo].LST_Adm1.Value & " | " & [F_Geo].LST_Adm2.Value & " | " & [F_Geo].LST_Adm3.Value & " | " & [F_Geo].LST_Adm4.Value)
    TXT_Msg.Value = sPlaceSelection
End Sub

Private Sub LST_AdmF1_Click()
    Call ShowLstF2(LST_AdmF1.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_AdmF2_Click()
    Call ShowLstF3(LST_AdmF2.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_AdmF3_Click()
    Call ShowLstF4(LST_AdmF3.Value)
    sPlaceSelection = TXT_Msg.Value
End Sub

Private Sub LST_AdmF4_Click()
    sPlaceSelection = ReverseString([F_Geo].LST_AdmF1.Value & " | " & [F_Geo].LST_AdmF2.Value & " | " & [F_Geo].LST_AdmF3.Value & " | " & [F_Geo].LST_AdmF4.Value)
    TXT_Msg.Value = sPlaceSelection

End Sub

'Those are trigerring event for the Histo
Private Sub LST_Histo_Click()
    TXT_Msg.Value = ReverseString(LST_Histo.Value)
    sPlaceSelection = LST_Histo.Value
End Sub

Private Sub LST_HistoF_Click()
    If LST_HistoF.Value <> "" Then
        TXT_Msg.Value = LST_HistoF.Value
        sPlaceSelection = LST_HistoF.Value
    End If
End Sub

Private Sub LST_ListeAgre_Click()
    TXT_Msg.Value = LST_ListeAgre.Value
    sPlaceSelection = LST_ListeAgre.Value
End Sub

Private Sub LST_ListeAgreF_Click()
    TXT_Msg.Value = LST_ListeAgreF.Value
    sPlaceSelection = LST_ListeAgreF.Value

End Sub

Private Sub TXT_Recherche_Change()
    'Search any value in geo data
    Call SearchValue(F_Geo.TXT_Recherche.Value)
End Sub

Private Sub TXT_RechercheF_Change()
    'Search any value in health facility
    Call SearchValueF(F_Geo.TXT_RechercheF.Value)

End Sub

Private Sub TXT_RechercheHisto_Change()
    'In case there is a change in the historic geographic Search list
    Call SeachHistoValue(F_Geo.TXT_RechercheHisto.Value)

End Sub

Private Sub TXT_RechercheHistoF_Change()
    'In case there is a change in the historic data
    Call SeachHistoValueF(F_Geo.TXT_RechercheHistoF.Value)

End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 650
    Me.height = 450

End Sub

