VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Geo 
   Caption         =   "GEO Apps"
   ClientHeight    =   9570.001
   ClientLeft      =   45
   ClientTop       =   -345
   ClientWidth     =   10200
   OleObjectBlob   =   "F_Geo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Geo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit

Const C_PWD As String = "1234"
Private Sub CMD_ChoicesFac_Click()

End Sub

Private Sub CMD_ChoicesGeo_Click()

End Sub

'This command is at the end, when you close the geoapp
'It basically update all the required data and input selected data in the linelist worksheet
Private Sub CMD_Copier_Click()

    Dim T_temp As BetterArray
    Dim geoSheet As String
    geoSheet = "GEO"
    Set T_temp = New BetterArray

    ActiveSheet.Unprotect (C_PWD)

    Select Case iGeoType
        'In case you selected the Geo data
    Case 0
        'updating the histo data if needed
        With Sheets(geoSheet).ListObjects("T_HistoGeo")
            If Not .DataBodyRange Is Nothing Then
                T_temp.FromExcelRange .DataBodyRange
                T_temp.Sort
                'only update if you don't find actual value then update
                If Not T_temp.Includes(TXT_Msg.value) Then
                    T_HistoGeo.Push ReverseString(TXT_Msg.value)
                End If
            Else
                'In case there is no histo data, update the first line
                If sPlaceSelection <> "" Then
                    T_HistoGeo.Push ReverseString(sPlaceSelection)
                End If
            End If
            'Now rewrite the histo data in the list object
            If T_HistoGeo.Length > 0 Then
                T_HistoGeo.Sort
                T_HistoGeo.ToExcelRange Destination:=Sheets(geoSheet).Range(Cells(2, .Range.Column).Address)
                'resize the list object
                .Resize .Range.CurrentRegion
                .DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlYes
            End If
        End With
        'Writing the selected data in the linelist sheet
        T_temp.Clear
        T_temp.Items = Split([TXT_Msg].value, " | ")
        If T_temp.Length > 0 Then
            Application.EnableEvents = False
            'Clear the cells before filling
            Range(ActiveCell.Address, ActiveCell.Offset(, 4)).value = ""
            T_temp.ToExcelRange Destination:=Range(ActiveCell.Address), TransposeValues:=True
            Application.EnableEvents = True
        End If
        T_temp.Clear
        'In Case we are dealing with the health facility (basically the same thing with little modifications)
    Case 1
        With Sheets(geoSheet).ListObjects("T_HistoHF")
            If Not .DataBodyRange Is Nothing Then
                T_temp.FromExcelRange .DataBodyRange
                T_temp.Sort

                If Not T_temp.Includes(ReverseString(TXT_Msg.value)) Then
                    T_HistoHF.Push [TXT_Msg].value
                End If
            Else
                If sPlaceSelection <> "" Then
                    T_HistoHF.Push sPlaceSelection
                End If
            End If
            'Now rewrite the histo data in the list object
            If (T_HistoHF.Length > 0) Then
                T_HistoHF.Sort
                T_HistoHF.ToExcelRange Destination:=Sheets(geoSheet).Range(Cells(2, .Range.Column).Address)
                'resize the list object
                .Resize .Range.CurrentRegion
                .DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlYes
            End If
        End With
        'writing the selected value
        Selection.value = TXT_Msg.value
    End Select
    [F_Geo].Hide
    'Protecting the worksheet
     ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
End Sub

'Closing the Geoapp
Private Sub CMD_Retour_Click()
    Me.Hide
End Sub

'Those are procedures to show the following list in one item is selected.
'They rely on ShowLst* functions coded in the Geo module
Private Sub LST_Adm1_Click()
    Call ShowLst2(LST_Adm1.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_Adm2_Click()
    Call ShowLst3(LST_Adm2.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_Adm3_Click()
    Call ShowLst4(LST_Adm3.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_Adm4_Click()
    sPlaceSelection = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value & " | " & [F_Geo].LST_Adm3.value & " | " & [F_Geo].LST_Adm4.value
    TXT_Msg.value = sPlaceSelection
End Sub

Private Sub LST_AdmF1_Click()
    Call ShowLstF2(LST_AdmF1.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_AdmF2_Click()
    Call ShowLstF3(LST_AdmF2.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_AdmF3_Click()
    Call ShowLstF4(LST_AdmF3.value)
    sPlaceSelection = TXT_Msg.value
End Sub

Private Sub LST_AdmF4_Click()
    sPlaceSelection = ReverseString([F_Geo].LST_AdmF1.value & " | " & [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF3.value & " | " & [F_Geo].LST_AdmF4.value)
    TXT_Msg.value = sPlaceSelection

End Sub

'Those are trigerring event for the Histo
Private Sub LST_Histo_Click()
    TXT_Msg.value = ReverseString(LST_Histo.value)
    sPlaceSelection = LST_Histo.value
End Sub

Private Sub LST_HistoF_Click()
    If LST_HistoF.value <> "" Then
        TXT_Msg.value = LST_HistoF.value
        sPlaceSelection = LST_HistoF.value
    End If
End Sub

Private Sub LST_ListeAgre_Click()
    TXT_Msg.value = LST_ListeAgre.value
    sPlaceSelection = LST_ListeAgre.value
End Sub

Private Sub LST_ListeAgreF_Click()
    TXT_Msg.value = LST_ListeAgreF.value
    sPlaceSelection = LST_ListeAgreF.value

End Sub

Private Sub TXT_Recherche_Change()
    'Search any value in geo data
    Call SearchValue(T_Concat, F_Geo.TXT_Recherche.value)
End Sub

Private Sub TXT_RechercheF_Change()
    'Search any value in health facility
    Call SearchValueF(T_ConcatHF, F_Geo.TXT_RechercheF.value)

End Sub

Private Sub TXT_RechercheHisto_Change()
    'In case there is a change in the historic geographic Search list
    Call SeachHistoValue(T_HistoGeo, F_Geo.TXT_RechercheHisto.value)

End Sub

Private Sub TXT_RechercheHistoF_Change()
    'In case there is a change in the historic data
    Call SeachHistoValueF(T_HistoHF, F_Geo.TXT_RechercheHistoF.value)

End Sub


