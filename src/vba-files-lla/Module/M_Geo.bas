Attribute VB_Name = "M_Geo"
Option Explicit
Option Base 1
'Module Geo: This is where the geo form and data are managed as well ass all geographic data
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'We keep the following data are public, so we can make some checks on their content even after initialization
'Geobases at country level
Public T_Adm1 As BetterArray                     'Administrative boundaries for level 1 (admin1) and 2 (admin2)
Public T_Adm2 As BetterArray                     'Administrative boundaries for level 2 (admin2) and 3 (admin3)
Public T_Adm3 As BetterArray                     'Administrative boundaries for level 3 (admin 3) and 4 (admin3)
Public T_Adm4 As BetterArray                     'Administrative boundaries table from in the geo database
Public T_HistoGeo As BetterArray                 'Historic of Geo
Public T_Concat As BetterArray                   'Binding everything for the concatenate for the Geo database

'Health facilities
Public T_HF As BetterArray                       'Health Facility data in the geo base
Public T_HF0 As BetterArray                      'Administrative boundaries for level 1 and 2 for Heath Facility
Public T_HF1 As BetterArray                      'administrative boundaries for level 2 and 3 for health facility
Public T_HF2 As BetterArray                      'Administrative boundaries for level 3 and 4 for health facility
Public T_HistoHF As BetterArray                  'Historic of health facility
Public T_ConcatHF   As BetterArray               'Health Facility concatenated

Public sPlaceSelection As String

'Those are number of rows of admin from 1 to 4. The goal is to use them to check if an update has been made in the geo sheet
Public iGeoType As Byte

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This sub loads the geodata from the Geo form to on form in the linenist. There are two types of data:
'Facility iGeoType = 1 or Geographical informations: iGeotype = 0 Some frame are hidden when one
'type of geodata should be load (facility) or (geographical informations)
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub LoadGeo(iGeoType As Byte)                    'Type of geo form to load: Geo = 0 or Facility = 1
    Dim geoSheet As String
    Dim transValue As Variant                    'transitional value for storing results
    geoSheet = "GEO"
    Dim i As Integer
    Dim T_HFTable As Variant
    
    Set T_Adm1 = New BetterArray
    Set T_Adm2 = New BetterArray
    Set T_Adm3 = New BetterArray
    Set T_Adm4 = New BetterArray
    Set T_Concat = New BetterArray
    Set T_HistoGeo = New BetterArray
    
    Set T_HF = New BetterArray
    Set T_ConcatHF = New BetterArray
    Set T_HistoHF = New BetterArray
    Set T_HF0 = New BetterArray
    Set T_HF1 = New BetterArray
    Set T_HF2 = New BetterArray

    Application.ScreenUpdating = False

    ' width and height of the geo formulaire
    [F_Geo].Height = 360
    [F_Geo].Width = 606

    'Before doing the whole all thing, we need to test if the T_Adm data is empty or not
    If (Not Sheets(geoSheet).ListObjects("T_ADM1").DataBodyRange Is Nothing) Then
        T_Adm1.FromExcelRange Sheets(geoSheet).ListObjects("T_ADM1").DataBodyRange
    End If
    
    If (Not Sheets(geoSheet).ListObjects("T_ADM2").DataBodyRange Is Nothing) Then
        T_Adm2.FromExcelRange Sheets(geoSheet).ListObjects("T_ADM2").DataBodyRange
    End If
    
    If (Not Sheets(geoSheet).ListObjects("T_ADM3").DataBodyRange Is Nothing) Then
        T_Adm3.FromExcelRange Sheets(geoSheet).ListObjects("T_ADM3").DataBodyRange
    End If
    
    If (Not Sheets(geoSheet).ListObjects("T_ADM4").DataBodyRange Is Nothing) Then
        T_Adm4.FromExcelRange Sheets(geoSheet).ListObjects("T_ADM4").DataBodyRange
    End If
       
        
    '----- Fill the list of the admins with the unique values for adm1
    [F_Geo].[LST_Adm1].List = T_Adm1.ExtractSegment(ColumnIndex:=1)
        
    '----- Add Caption for  each adminstrative leveles in the form
    F_Geo.LBL_Adm1.Caption = Sheets(geoSheet).ListObjects("T_ADM4").HeaderRowRange.Item(1).value
    F_Geo.LBL_Adm2.Caption = Sheets(geoSheet).ListObjects("T_ADM4").HeaderRowRange.Item(2).value
    F_Geo.LBL_Adm3.Caption = Sheets(geoSheet).ListObjects("T_ADM4").HeaderRowRange.Item(3).value
    F_Geo.LBL_Adm4.Caption = Sheets(geoSheet).ListObjects("T_ADM4").HeaderRowRange.Item(4).value
        
    '------- Concatenate all the tables for the geo
    For i = T_Adm4.LowerBound To T_Adm4.UpperBound
        'binding all the lines together
        transValue = T_Adm4.Item(i)              'This is oneline of the adm
        T_Concat.Item(i) = CStr(transValue(1)) & " | " & CStr(transValue(2)) & " | " & CStr(transValue(3)) & " | " & CStr(transValue(4))
    Next
    T_Concat.Sort
    '------ Once the concat is created, add it to the list in the form
    [F_Geo].LST_ListeAgre.List = T_Concat.Items
    
    ' Now health facility ----------------------------------------------------------------------------------------------------------
    If (Not Sheets(geoSheet).ListObjects("T_HF").DataBodyRange Is Nothing) Then
       
        T_HFTable = Sheets(geoSheet).ListObjects("T_HF").DataBodyRange
        T_HF.Items = T_HFTable
      
        'unique admin 1
        T_HF0.Items = GetUnique(T_HFTable, 4)
        T_HF1.Items = GetUnique(T_HFTable, 4, 3)
        T_HF2.Items = GetUnique(T_table:=T_HFTable, index:=Array(4, 3, 2))
         
        ReDim T_HFTable(1)
                         
        ' ----- Fill the list of the admins with the unique values of adm1
        [F_Geo].[LST_AdmF1].List = T_HF0.Items
                
        '-------- Adding caption for each admnistrative levels in the form of the health facility
        F_Geo.LBL_Adm1F.Caption = Sheets(geoSheet).ListObjects("T_HF").HeaderRowRange.Item(4).value
        F_Geo.LBL_Adm2F.Caption = Sheets(geoSheet).ListObjects("T_HF").HeaderRowRange.Item(3).value
        F_Geo.LBL_Adm3F.Caption = Sheets(geoSheet).ListObjects("T_HF").HeaderRowRange.Item(2).value
        F_Geo.LBL_Adm4F.Caption = Sheets(geoSheet).ListObjects("T_HF").HeaderRowRange.Item(1).value
        
        'Creating the concatenate for the Health facility
        For i = T_HF.LowerBound To T_HF.UpperBound
            transValue = T_HF.Item(i)
            T_ConcatHF.Item(i) = CStr(transValue(1)) & " | " & CStr(transValue(2)) & " | " & CStr(transValue(3)) & " | " & CStr(transValue(4))
        Next i
        T_ConcatHF.Sort
        '---- Once the concat is created, add it to the HF form using the list for the concat part
        [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
    End If
    'Historic for geographic data and facility data
    If Not Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange Is Nothing Then
        T_HistoGeo.FromExcelRange Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange
        [F_Geo].LST_Histo.List = T_HistoGeo.Items
    End If

    If Not Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange Is Nothing Then
        T_HistoHF.FromExcelRange Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange
        [F_Geo].LST_HistoF.List = T_HistoHF.Items
    End If

    'Showing the form in case of Geo or Health Facility. Geo and Facility are in different frames.
    Select Case iGeoType
    Case 0
        [F_Geo].FRM_Facility.Visible = False
        [F_Geo].FRM_Geo.Visible = True
        [F_Geo].LBL_Fac1.Visible = False
        [F_Geo].LBL_Geo1.Visible = True
    Case 1
        [F_Geo].FRM_Facility.Visible = True
        [F_Geo].FRM_Geo.Visible = False
        [F_Geo].LBL_Fac1.Visible = True
        [F_Geo].LBL_Geo1.Visible = False
    End Select
    Application.ScreenUpdating = True
    
    [F_Geo].show
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'This sub shows the list of the selected values in the geo frame (second list) given a place sPlace
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This function shows the second list for the Geo
Sub ShowLst2(sPlace As String)
    'clear the forms if there is something
    [F_Geo].LST_Adm2.Clear
    [F_Geo].LST_Adm3.Clear
    [F_Geo].LST_Adm4.Clear
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list

    Set T_Aff = New BetterArray
    
    'Search if the value exists in the 2 dimensional table T_Adm1 previously initialized
    If T_Adm2.Length > 0 Then
        Set T_Aff = GetFilter(T_Adm2, 1, sPlace)
    End If
    
    [F_Geo].TXT_Msg.value = sPlace
    'update if only next level is available
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm2.List = T_Aff.ExtractSegment(, ColumnIndex:=2)
    End If
    'Clear
    T_Aff.Clear
End Sub

'Show second list for the facility
Sub ShowLstF2(sPlace As String)

    'Clear the forms
    [F_Geo].LST_AdmF2.Clear
    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff = New BetterArray

    If T_HF1.Length > 0 Then
        Set T_Aff = GetFilter(T_HF1, 1, sPlace)
    End If
    
    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF2.List = T_Aff.ExtractSegment(, ColumnIndex:=2)
        [F_Geo].TXT_Msg.value = sPlace
    Else
        [F_Geo].TXT_Msg.value = sPlace           'No levels
    End If
    T_Aff.Clear
End Sub

'This function shows the third list for the geobase
Sub ShowLst3(sPlace As String)
     
    'Clear the two remaining forms
    [F_Geo].LST_Adm3.Clear
    [F_Geo].LST_Adm4.Clear
    Dim sAdm1 As String 'Selected admin 1
    
    Dim T_Aff1 As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff1 = New BetterArray
    
    Dim T_Aff As BetterArray
    Set T_Aff = New BetterArray
    
    sAdm1 = [F_Geo].LST_Adm1.value
    
    If T_Adm3.Length > 0 Then
        'Filter on the adm 1 firts
        Set T_Aff1 = GetFilter(T_Adm3, 1, sAdm1)
        'Then filter on adm2
        Set T_Aff = GetFilter(T_Aff1, 2, sPlace)
        T_Aff1.Clear
        Set T_Aff1 = Nothing
    End If
    
    'Update the adm3 list in the geoform if the T_Aff3 is not missing
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm3.List = T_Aff.ExtractSegment(, ColumnIndex:=3)
        [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value
    Else
        [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value 'No lower level related
    End If
    T_Aff.Clear
End Sub

'Show the third list of geobase, pretty much the same as before
Sub ShowLstF3(sPlace As String)

    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear
    
    'first level
    Dim T_Aff1 As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff1 = New BetterArray
    
    'second and last level
    Dim T_Aff As BetterArray
    Set T_Aff = New BetterArray
    
    Dim sAdm As String
    sAdm = [F_Geo].LST_AdmF1.value
    
    If T_HF2.Length > 0 Then
        'Filter on Adm1
        Set T_Aff1 = GetFilter(T_HF2, 1, sAdm)
        'Then filter on Adm2
        Set T_Aff = GetFilter(T_Aff1, 2, sPlace)
        T_Aff1.Clear
        Set T_Aff1 = Nothing
    End If
    
    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF3.List = T_Aff.ExtractSegment(, ColumnIndex:=3)
        [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value
    Else
        [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value 'No level 3
    End If
    
    T_Aff.Clear
End Sub

'This function shows the fourth list for the Geo (pretty much the same thing as done previously)
Sub ShowLst4(sPlace As String)

    [F_Geo].LST_Adm4.Clear
    
    Dim T_Aff1 As BetterArray
    Set T_Aff1 = New BetterArray
    
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff = New BetterArray
    
    Dim sAdm As String         'Adms selected previously
    sAdm = [F_Geo].LST_Adm1.value
    If T_Adm4.Length > 0 Then
    'Filter on adm1
        Set T_Aff = GetFilter(T_Adm4, 1, sAdm)
        sAdm = [F_Geo].LST_Adm2.value
        'Then filter result on adm2
        Set T_Aff1 = GetFilter(T_Aff, 2, sAdm)
        T_Aff.Clear
        
        'Then filter on selected adm3
        Set T_Aff = GetFilter(T_Aff1, 3, sPlace)
        T_Aff1.Clear
        Set T_Aff1 = Nothing
    End If
    
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm4.List = T_Aff.ExtractSegment(, ColumnIndex:=4)
        [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value & " | " & [F_Geo].LST_Adm3.value
    Else
        [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value & " | " & [F_Geo].LST_Adm3.value 'No level found
    End If
    T_Aff.Clear
End Sub

'Fourth list of health facility
Sub ShowLstF4(sPlace As String)

    [F_Geo].LST_AdmF4.Clear
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff = New BetterArray
    
    Dim T_Aff1 As BetterArray
    Set T_Aff1 = New BetterArray
    
    Dim sAdm As String 'previously selected admin levels
    sAdm = [F_Geo].LST_AdmF1.value
    
    If T_HF.Length > 0 Then
        'Filter on adm1
        Set T_Aff = GetFilter(T_HF, 4, sAdm)
        'Filter on adm2
        sAdm = [F_Geo].LST_AdmF2.value
        Set T_Aff1 = GetFilter(T_Aff, 3, sAdm)
        T_Aff.Clear
        'now on adm3
        sAdm = [F_Geo].LST_AdmF3.value
        Set T_Aff = GetFilter(T_Aff1, 2, sAdm)
        T_Aff1.Clear
        Set T_Aff1 = Nothing
    End If

    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF4.List = T_Aff.ExtractSegment(, ColumnIndex:=1)
        [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF3.value & " | " & [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value
    Else
        [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF3.value & " | " & [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value '& " : Pas de niveau 4"
    End If
End Sub

Sub ClearGeo()
    ' Clear all elements in the geo form
    [F_Geo].LST_Adm1.Clear
    [F_Geo].LST_Adm2.Clear
    [F_Geo].LST_Adm3.Clear
    [F_Geo].LST_Adm4.Clear
    [F_Geo].LST_ListeAgre.Clear
    [F_Geo].LST_AdmF1.Clear
    [F_Geo].LST_AdmF2.Clear
    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear
    [F_Geo].LST_ListeAgreF.Clear
    [F_Geo].TXT_Msg.value = ""
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Search values in the search box
'T_Concat is The concatenate data
'Search value is the string to search from Those type of functions are pretty much the same:
'1 - search in the concatenated table
'2- Add values where there are some matches in another table
'3- Render the table if it is not empty
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub SearchValue(T_Concat, ByVal sSearchedValue As String)
    Dim T_result As BetterArray
    Set T_result = New BetterArray
    Dim i As Integer
    
    'Create a table of the found values (called T_result)
    If Len(sSearchedValue) >= 3 Then
        i = 1
        While i <= T_Concat.UpperBound
            If InStr(1, LCase(T_Concat.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_Concat.Item(i)
            End If
            i = i + 1
        Wend
        
        'Render the table if some values are found
        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_ListeAgre.List = T_result.Items
        Else
            'If Not, check if there have been some input in the concat and render
            If [F_Geo].LST_ListeAgre.ListCount - 1 <> T_Concat.UpperBound Then
                [F_Geo].LST_ListeAgre.List = T_Concat.Items
            End If
        End If
    Else
        If [F_Geo].LST_ListeAgre.ListCount - 1 <> T_Concat.UpperBound Then
            [F_Geo].LST_ListeAgre.List = T_Concat.Items
        End If
    End If
    
    Set T_result = Nothing
End Sub

Sub SeachHistoValue(T_HistoGeo, sSearchedValue As String)
    Dim T_result As BetterArray
    Dim i As Integer
    
    If Len(sSearchedValue) >= 3 Then
        Set T_result = New BetterArray
        i = 1
        While i <= T_HistoGeo.UpperBound
            If InStr(1, LCase(T_HistoGeo.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_HistoGeo.Item(i)
            End If
            i = i + 1
        Wend
    
        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_Histo.List = T_result.Items
        Else
            If [F_Geo].LST_Histo.ListCount - 1 <> T_HistoGeo.UpperBound Then
                [F_Geo].LST_Histo.List = T_HistoGeo.Items
            End If
        End If
    Else
        If [F_Geo].LST_Histo.ListCount - 1 <> T_HistoGeo.UpperBound Then
            [F_Geo].LST_Histo.List = T_HistoGeo.Items
        End If
    End If

    Set T_result = Nothing
End Sub

Sub SearchValueF(T_ConcatHF, sSearchedValue As String)
    Dim T_result As BetterArray
    Dim i As Integer

    If Len(sSearchedValue) >= 3 Then
        Set T_result = New BetterArray
        i = 1
        While i <= T_ConcatHF.UpperBound
            If InStr(1, LCase(T_ConcatHF.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_ConcatHF.Item(i)
            End If
            i = i + 1
        Wend
    
        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_ListeAgreF.List = T_result.Items
        Else
            If [F_Geo].LST_ListeAgreF.ListCount - 1 <> T_ConcatHF.UpperBound Then
                [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
            End If
        End If
    Else
        If [F_Geo].LST_ListeAgreF.ListCount - 1 <> T_ConcatHF.UpperBound Then
            [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
        End If
    End If

    Set T_result = Nothing
End Sub

Sub SeachHistoValueF(T_HistoHF, sSearchedValue As String)

    Dim T_result As BetterArray
    Dim i As Integer

    If Len(sSearchedValue) >= 3 Then
        i = 1
        Set T_result = New BetterArray
        
        While i <= T_HistoHF.UpperBound
            If InStr(1, LCase(T_HistoHF.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_HistoHF.Item(i)
            End If
            i = i + 1
        Wend
    
        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_HistoF.List = T_result.Items
        Else
            If [F_Geo].LST_HistoF.ListCount - 1 <> T_HistoHF.UpperBound Then
                [F_Geo].LST_HistoF.List = T_HistoHF.Items
            End If
        End If
    Else
        If [F_Geo].LST_HistoF.ListCount - 1 <> T_HistoHF.UpperBound Then
            [F_Geo].LST_HistoF.List = T_HistoHF.Items
        End If
    End If

    Set T_result = Nothing
End Sub

' This function reverses a string using the | as separator, like in the final selection of the
' Health facility form.
Function ReverseString(sChaine As String)
    Dim i As Integer
    Dim T_temp As BetterArray
    Set T_temp = New BetterArray
    T_temp.LowerBound = 1
    Dim sRes As String
    
    ReverseString = ""
    T_temp.Items = Split(sChaine, " | ")
    sRes = T_temp.Items(1)
    For i = T_temp.LowerBound + 1 To T_temp.UpperBound
        sRes = T_temp.Items(i) & " | " & sRes
    Next
     
    ReverseString = sRes
    Set T_temp = Nothing
End Function



