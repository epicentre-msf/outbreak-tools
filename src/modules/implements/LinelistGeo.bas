Attribute VB_Name = "LinelistGeo"

Option Explicit
Option Base 1
Option Private Module


Public T_HistoGeo As BetterArray                 'Historic of Geo
Public sPlaceSelection As String
Public T_HistoHF As BetterArray                  'Historic of health facility

'Health facilities
Private T_Concat As BetterArray                   'Binding everything for the concatenate for the Geo database
Private T_ConcatHF   As BetterArray               'Health Facility concatenated
Private geo As ILLGeo 'The geo object is used in the entire module for filtering, for sorting and also for computing various things


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This sub loads the geodata from the Geo form to on form in the linenist. There are two types of data:
'Facility iGeoType = 1 or Geographical informations: iGeotype = 0 Some frame are hidden when one
'type of geodata should be load (facility) or (geographical informations)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub LoadGeo(iGeoType As Byte)                    'Type of geo form to load: Geo = 0 or Facility = 1
    Dim sh As Worksheet
    Dim transValue As BetterArray

    Set T_Concat = New BetterArray
    Set T_HistoGeo = New BetterArray
    Set transValue = New BetterArray

    Set T_ConcatHF = New BetterArray
    Set T_HistoHF = New BetterArray
    Set transValue = New BetterArray
    Set sh = ThisWorkbook.Worksheets("Geo")

    On Error GoTo ErrLoadGeo
    Set geo = LLGeo.Create(sh)

    BeginWork xlsapp:=Application

    With sh

        Select Case iGeoType
            'Load Geo informations
        Case 0
            'Add Caption for  each adminstrative leveles in the form
            F_Geo.LBL_Adm1.Caption = .ListObjects("T_ADM4").HeaderRowRange.Item(1).Value
            F_Geo.LBL_Adm2.Caption = .ListObjects("T_ADM4").HeaderRowRange.Item(2).Value
            F_Geo.LBL_Adm3.Caption = .ListObjects("T_ADM4").HeaderRowRange.Item(3).Value
            F_Geo.LBL_Adm4.Caption = .ListObjects("T_ADM4").HeaderRowRange.Item(4).Value

            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin4")
            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin3")
            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin2")

            'Before doing the whole all thing, we need to test if the T_Adm data is empty or not

            If (Not .ListObjects("T_ADM4").DataBodyRange Is Nothing) Then

                'T_Adm4.FromExcelRange .ListObjects("T_ADM4").DataBodyRange
                Set transValue = geo.GeoLevel(LevelAdmin1, GeoScopeAdmin)
                [F_Geo].[LST_Adm1].List = transValue.Items
                T_Concat.FromExcelRange .ListObjects("T_ADM4").ListColumns("adm4_concat").DataBodyRange
                T_Concat.Sort
                '------ Once the concat is created, add it to the list in the form
                [F_Geo].LST_ListeAgre.List = T_Concat.Items
            End If

            'Historic for geographic data and facility data
            If Not .ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
                T_HistoGeo.FromExcelRange .ListObjects(C_sTabHistoGeo).DataBodyRange
                [F_Geo].LST_Histo.List = T_HistoGeo.Items
            End If

            [F_Geo].FRM_Facility.Visible = False
            [F_Geo].FRM_Geo.Visible = True
            [F_Geo].LBL_Fac1.Visible = False
            [F_Geo].LBL_Geo1.Visible = True

        Case 1
            'Adding caption for each admnistrative levels in the form of the health facility
            F_Geo.LBL_Adm1F.Caption = .ListObjects("T_HF").HeaderRowRange.Item(4).Value
            F_Geo.LBL_Adm2F.Caption = .ListObjects("T_HF").HeaderRowRange.Item(3).Value
            F_Geo.LBL_Adm3F.Caption = .ListObjects("T_HF").HeaderRowRange.Item(2).Value
            F_Geo.LBL_Adm4F.Caption = .ListObjects("T_HF").HeaderRowRange.Item(1).Value

            'Now health facility ----------------------------------------------------------------------------------------------------------
            If (Not .ListObjects("T_HF").DataBodyRange Is Nothing) Then

                transValue.Clear
                'unique admin 1
                transValue.FromExcelRange .ListObjects("T_HF").ListColumns(4).DataBodyRange
                Set transValue = GetUniqueBA(transValue)
                ' ----- Fill the list of the admins with the unique values of adm1
                [F_Geo].[LST_AdmF1].List = transValue.Items

                T_ConcatHF.FromExcelRange .ListObjects("T_HF").ListColumns("hf_concat").DataBodyRange
                T_ConcatHF.Sort
                '---- Once the concat is created, add it to the HF form using the list for the concat part
                [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
            End If

            'Historic HF
            If Not .ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
                T_HistoHF.FromExcelRange .ListObjects(C_sTabHistoHF).DataBodyRange
                [F_Geo].LST_HistoF.List = T_HistoHF.Items
            End If

            [F_Geo].FRM_Facility.Visible = True
            [F_Geo].FRM_Geo.Visible = False
            [F_Geo].LBL_Fac1.Visible = True
            [F_Geo].LBL_Geo1.Visible = False

        End Select
    End With


    EndWork xlsapp:=Application

    [F_Geo].TXT_Msg.Value = vbNullString
    [F_Geo].Show

    Exit Sub

ErrLoadGeo:

    MsgBox TranslateLLMsg("MSG_ErrGeo"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
    EndWork xlsapp:=Application

    Exit Sub
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
    Dim T_Aff As BetterArray

    Set T_Aff = New BetterArray
    Application.Cursor = xlNorthwestArrow

    'Search if the value exists in the 2 dimensional table T_Adm1 previously initialized
    Set T_Aff = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, sPlace)

    [F_Geo].TXT_Msg.Value = sPlace
    'update if only next level is available
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm2.List = T_Aff.Items
    End If

    Application.Cursor = xlDefault
End Sub

'Show second list for the facility
Sub ShowLstF2(sPlace As String)

    'Clear the forms
    [F_Geo].LST_AdmF2.Clear
    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list
    Set T_Aff = New BetterArray

    Application.Cursor = xlNorthwestArrow

    'Just filter and show
    Set T_Aff = geo.GeoLevel(LevelAdmin2, GeoScopeHF, sPlace)

    [F_Geo].TXT_Msg.Value = sPlace

    If T_Aff.Length > 0 Then [F_Geo].LST_AdmF2.List = T_Aff.Items

    Application.Cursor = xlDefault

End Sub

'This function shows the third list for the geobase
Sub ShowLst3(sAdm2 As String)

    'Clear the two remaining forms
    [F_Geo].LST_Adm3.Clear
    [F_Geo].LST_Adm4.Clear

    Dim sAdm1 As String                          'Selected admin 1
    Dim T_Aff As BetterArray
    Dim adminNames As BetterArray

    Set adminNames = New BetterArray
    adminNames.LowerBound = 1 'It is important to have 1 as lowerbound for the filterig
    sAdm1 = [F_Geo].LST_Adm1.Value

    Application.Cursor = xlNorthwestArrow

    'add admin 1 and admin 2 for filtering
    adminNames.Push sAdm1, sAdm2

    'Just filter and show
    Set T_Aff = geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, adminNames)

    [F_Geo].TXT_Msg.Value = [F_Geo].LST_Adm1.Value & " | " & [F_Geo].LST_Adm2.Value
    'Update the adm3 list in the geoform if the T_Aff3 is not missing
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm3.List = T_Aff.Items
    End If

    Application.Cursor = xlDefault

End Sub

'Show the third list of geobase, pretty much the same as before
Sub ShowLstF3(sAdm2 As String)

    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear

    Dim sAdm1 As String
    Dim T_Aff As BetterArray
    Dim adminNames As BetterArray

    sAdm1 = [F_Geo].LST_AdmF1.Value
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1

    adminNames.Push sAdm1, sAdm2 'selected admin 1 and 2

    Application.Cursor = xlNorthwestArrow

    'Here I filter on the Health Facilities instead of the Geo
    Set T_Aff = geo.GeoLevel(LevelAdmin3, GeoScopeHF, adminNames)

    [F_Geo].TXT_Msg.Value = [F_Geo].LST_AdmF2.Value & " | " & [F_Geo].LST_AdmF1.Value

    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF3.List = T_Aff.Items
    End If

    Application.Cursor = xlDefault

End Sub

'This function shows the fourth list for the Geo (pretty much the same thing as done previously)
Sub ShowLst4(sAdm3 As String)

    [F_Geo].LST_Adm4.Clear

    Dim T_Aff As BetterArray
    Dim sAdm1 As String
    Dim sAdm2 As String
    Dim adminNames As BetterArray

    sAdm1 = [F_Geo].LST_Adm1.Value
    sAdm2 = [F_Geo].LST_Adm2.Value
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1
    adminNames.Push sAdm1, sAdm2, sAdm3 'Add the three levels to filter up to level 4

    Application.Cursor = xlNorthwestArrow

    [F_Geo].TXT_Msg.Value = [F_Geo].LST_Adm1.Value & " | " & [F_Geo].LST_Adm2.Value & " | " & [F_Geo].LST_Adm3.Value

    Set T_Aff = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)

    If T_Aff.Length > 0 Then [F_Geo].LST_Adm4.List = T_Aff.Items

    Application.Cursor = xlDefault
End Sub

'Fourth list of health facility
Sub ShowLstF4(sAdm3 As String)

    [F_Geo].LST_AdmF4.Clear

    Dim T_Aff As BetterArray
    Dim adminNames As BetterArray
    Dim sAdm1 As String
    Dim sAdm2 As String

    sAdm1 = [F_Geo].LST_AdmF1.Value
    sAdm2 = [F_Geo].LST_AdmF2.Value
    Set adminNames = New BetterArray

    adminNames.LowerBound = 1
    adminNames.Push sAdm1, sAdm2, sAdm3

    'Use the cursor to hide some working steps
    Application.Cursor = xlNorthwestArrow

    Set T_Aff = geo.GeoLevel(LevelAdmin4, GeoScopeHF, adminNames)

    [F_Geo].TXT_Msg.Value = [F_Geo].LST_AdmF3.Value & " | " & [F_Geo].LST_AdmF2.Value & " | " & [F_Geo].LST_AdmF1.Value

    If T_Aff.Length > 0 Then [F_Geo].LST_AdmF4.List = T_Aff.Items

    Application.Cursor = xlDefault
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
    [F_Geo].TXT_Msg.Value = ""
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Search values in the search box
'T_Concat is The concatenate data
'Search value is the string to search from Those type of functions are pretty much the same:
'1 - search in the concatenated table
'2- Add values where there are some matches in another table
'3- Render the table if it is not empty
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub SearchValue(ByVal sSearchedValue As String)
    Dim T_result As BetterArray
    Set T_result = New BetterArray
    Dim i As Long



    'Create a table of the found values (called T_result)
    If Len(sSearchedValue) >= 3 Then
        i = 1
        Do While i <= T_Concat.UpperBound
            If InStr(1, LCase(T_Concat.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_Concat.Item(i)
            End If
            i = i + 1
        Loop

        'Render the table if some values are found
        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_ListeAgre.List = T_result.Items
        Else
            'If Not, check if there have been some input in the concat and render
            [F_Geo].LST_ListeAgre.Clear
        End If
    Else
        If [F_Geo].LST_ListeAgre.ListCount - 1 <> T_Concat.UpperBound Then
            [F_Geo].LST_ListeAgre.List = T_Concat.Items
        End If
    End If

End Sub

Sub SeachHistoValue(sSearchedValue As String)
    Dim T_result As BetterArray
    Dim i As Long

    If Len(sSearchedValue) >= 3 Then
        Set T_result = New BetterArray
        i = 1
        Do While i <= T_HistoGeo.UpperBound
            If InStr(1, LCase(T_HistoGeo.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_HistoGeo.Item(i)
            End If
            i = i + 1
        Loop

        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_Histo.List = T_result.Items
        Else
            [F_Geo].LST_Histo.Clear
        End If
    Else
        If [F_Geo].LST_Histo.ListCount - 1 <> T_HistoGeo.UpperBound Then
            [F_Geo].LST_Histo.List = T_HistoGeo.Items
        End If
    End If

End Sub

Sub SearchValueF(sSearchedValue As String)
    Dim T_result As BetterArray
    Dim i As Long

    If Len(sSearchedValue) >= 3 Then
        Set T_result = New BetterArray
        i = 1
        Do While i <= T_ConcatHF.UpperBound
            If InStr(1, LCase(T_ConcatHF.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_ConcatHF.Item(i)
            End If
            i = i + 1
        Loop

        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_ListeAgreF.List = T_result.Items
        Else
            [F_Geo].LST_ListeAgreF.Clear
        End If
    Else
        If [F_Geo].LST_ListeAgreF.ListCount - 1 <> T_ConcatHF.UpperBound Then
            [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
        End If
    End If

End Sub

Sub SeachHistoValueF(sSearchedValue As String)

    Dim T_result As BetterArray
    Dim i As Long

    If Len(sSearchedValue) >= 3 Then
        i = 1
        Set T_result = New BetterArray

        Do While i <= T_HistoHF.UpperBound
            If InStr(1, LCase(T_HistoHF.Item(i)), LCase(sSearchedValue)) > 0 Then
                T_result.Push T_HistoHF.Item(i)
            End If
            i = i + 1
        Loop

        If T_result.Length > 0 Then
            T_result.Sort
            [F_Geo].LST_HistoF.List = T_result.Items
        Else
            [F_Geo].LST_HistoF.Clear
        End If
    Else
        If [F_Geo].LST_HistoF.ListCount - 1 <> T_HistoHF.UpperBound Then
            [F_Geo].LST_HistoF.List = T_HistoHF.Items
        End If
    End If

End Sub

' This function reverses a string using the | as separator, like in the final selection of the
' Health facility form.
Function ReverseString(sChaine As String)
    Dim i As Long
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
End Function

Sub ClearOneHistoricGeobase(iGeoType As Byte)
    Dim WkshGeo As Worksheet
    Dim ShouldDelete As Long

    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    ShouldDelete = MsgBox("Your historic geographic data  in the current workbook will be completely deleted, this action is irreversible. Proceed?", vbExclamation + vbYesNo, "Delete Historic")

    If ShouldDelete = vbYes Then
        If iGeoType = 0 Then
            If Not WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange.Delete
                T_HistoGeo.Clear
                [F_Geo].LST_Histo.Clear
            End If
        End If
        If iGeoType = 1 Then
            If Not WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange.Delete
                T_HistoHF.Clear
                [F_Geo].LST_HistoF.Clear
            End If
        End If

        MsgBox "Done", vbInformation, "Delete Historic"
    End If
End Sub

