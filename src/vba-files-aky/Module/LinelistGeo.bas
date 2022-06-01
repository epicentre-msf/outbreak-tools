Attribute VB_Name = "LinelistGeo"

Option Explicit
Option Base 1
'Module Geo: This is where the geo form and data are managed as well ass all geographic data
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'We keep the following data are public, so we can make some checks on their content even after initialization
'Geobases at country level
Public T_Adm4 As BetterArray                     'Administrative boundaries for level 1 (admin1) and 2 (admin2)
Public T_HistoGeo As BetterArray                 'Historic of Geo
Public T_Concat As BetterArray                   'Binding everything for the concatenate for the Geo database

'Health facilities
Public T_HF As BetterArray                       'Health Facility data in the geo base
Public T_HistoHF As BetterArray                  'Historic of health facility
Public T_ConcatHF   As BetterArray               'Health Facility concatenated

Public sPlaceSelection As String


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This sub loads the geodata from the Geo form to on form in the linenist. There are two types of data:
'Facility iGeoType = 1 or Geographical informations: iGeotype = 0 Some frame are hidden when one
'type of geodata should be load (facility) or (geographical informations)
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub LoadGeo(iGeoType As Byte)                    'Type of geo form to load: Geo = 0 or Facility = 1
    Dim transValue As BetterArray
    Dim i As Integer
    Set T_Adm4 = New BetterArray
    Set T_Concat = New BetterArray
    Set T_HistoGeo = New BetterArray
    Set transValue = New BetterArray

    Set T_HF = New BetterArray
    Set T_ConcatHF = New BetterArray
    Set T_HistoHF = New BetterArray
    Set transValue = New BetterArray

    On Error GoTo ErrLoadGeo

    BeginWork xlsapp:=Application
    With ThisWorkbook.Worksheets(C_sSheetGeo)

        Select Case iGeoType

            'Load Geo informations
            Case 0

            'Add Caption for  each adminstrative leveles in the form
            F_Geo.LBL_Adm1.Caption = .ListObjects(C_sTabAdm4).HeaderRowRange.Item(1).value
            F_Geo.LBL_Adm2.Caption = .ListObjects(C_sTabAdm4).HeaderRowRange.Item(2).value
            F_Geo.LBL_Adm3.Caption = .ListObjects(C_sTabAdm4).HeaderRowRange.Item(3).value
            F_Geo.LBL_Adm4.Caption = .ListObjects(C_sTabAdm4).HeaderRowRange.Item(4).value


            'Before doing the whole all thing, we need to test if the T_Adm data is empty or not
            If (Not .ListObjects(C_sTabAdm4).DataBodyRange Is Nothing) Then
                T_Adm4.FromExcelRange .ListObjects(C_sTabAdm4).DataBodyRange
                transValue.FromExcelRange .ListObjects(C_sTabAdm4).ListColumns(1).DataBodyRange
                transValue.Sort
                Set transValue = GetUniqueBA(transValue)
                [F_Geo].[LST_Adm1].List = transValue.Items
                '------- Concatenate all the tables for the geo
                For i = T_Adm4.LowerBound To T_Adm4.UpperBound
                    transValue.Clear
                    'binding all the lines together
                    transValue.Items = T_Adm4.Item(i)              'This is oneline of the adm
                    T_Concat.Item(i) = transValue.Item(1) & " | " & transValue.Item(2) & " | " & transValue.Item(3) & " | " & transValue.Item(4)
                Next
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

            'Load Health Facility

            Case 1
            '-------- Adding caption for each admnistrative levels in the form of the health facility
            F_Geo.LBL_Adm1F.Caption = .ListObjects(C_sTabHF).HeaderRowRange.Item(4).value
            F_Geo.LBL_Adm2F.Caption = .ListObjects(C_sTabHF).HeaderRowRange.Item(3).value
            F_Geo.LBL_Adm3F.Caption = .ListObjects(C_sTabHF).HeaderRowRange.Item(2).value
            F_Geo.LBL_Adm4F.Caption = .ListObjects(C_sTabHF).HeaderRowRange.Item(1).value

            'Now health facility ----------------------------------------------------------------------------------------------------------
            If (Not .ListObjects(C_sTabHF).DataBodyRange Is Nothing) Then

                T_HF.FromExcelRange .ListObjects(C_sTabHF).DataBodyRange
                transValue.Clear
                'unique admin 1
                transValue.FromExcelRange .ListObjects(C_sTabHF).ListColumns(4).DataBodyRange
                Set transValue = GetUniqueBA(transValue)
                ' ----- Fill the list of the admins with the unique values of adm1
                [F_Geo].[LST_AdmF1].List = transValue.Items
                'Creating the concatenate for the Health facility
                For i = T_HF.LowerBound To T_HF.UpperBound
                    transValue.Clear
                    transValue.Items = T_HF.Item(i)
                    T_ConcatHF.Item(i) = transValue.ToString(Separator:="|", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                Next i
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

    Set transValue = Nothing
    [F_Geo].TXT_Msg.value = vbNullString
    [F_Geo].show

    Exit Sub

    ErrLoadGeo:
        MsgBox "Error, Unable to Load the GeoApp"
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
    Dim Wksh As Worksheet                     'Aff is for rendering filtered values withing the list

    Set T_Aff = New BetterArray
    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)

    'Search if the value exists in the 2 dimensional table T_Adm1 previously initialized
    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabAdm2), 1, sPlace, returnIndex:=2)

    [F_Geo].TXT_Msg.value = sPlace
    'update if only next level is available
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm2.List = T_Aff.Items
    End If

    'Clear
    Set T_Aff = Nothing
    Set Wksh = Nothing
End Sub

'Show second list for the facility
Sub ShowLstF2(sPlace As String)

    'Clear the forms
    [F_Geo].LST_AdmF2.Clear
    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear
    Dim T_Aff As BetterArray                     'Aff is for rendering filtered values withing the list
    Dim Wksh As Worksheet

    Set T_Aff = New BetterArray
    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)

    'Just filter and show
    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabHF), 4, sPlace, returnIndex:=3)
    Set T_Aff = GetUniqueBA(T_Aff)

    [F_Geo].TXT_Msg.value = sPlace

    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF2.List = T_Aff.Items
    End If

    Set T_Aff = Nothing
    Set Wksh = Nothing
End Sub

'This function shows the third list for the geobase
Sub ShowLst3(sAdm2 As String)

    'Clear the two remaining forms
    [F_Geo].LST_Adm3.Clear
    [F_Geo].LST_Adm4.Clear

    Dim sAdm1 As String 'Selected admin 1
    Dim T_Aff As BetterArray
    Dim Wksh As Worksheet

    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)
    sAdm1 = [F_Geo].LST_Adm1.value

    'Just filter and show
    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabAdm3), 1, sAdm1, 2, sAdm2, returnIndex:=3)

    [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value
    'Update the adm3 list in the geoform if the T_Aff3 is not missing
    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm3.List = T_Aff.Items
    End If

    Set T_Aff = Nothing
    Set Wksh = Nothing
End Sub

'Show the third list of geobase, pretty much the same as before
Sub ShowLstF3(sAdm2 As String)

    [F_Geo].LST_AdmF3.Clear
    [F_Geo].LST_AdmF4.Clear

    Dim sAdm1 As String
    Dim T_Aff As BetterArray
    Dim Wksh As Worksheet

    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)
    sAdm1 = [F_Geo].LST_AdmF1.value

    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabHF), 4, sAdm1, 3, sAdm2, returnIndex:=2)

    [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value
    Set T_Aff = GetUniqueBA(T_Aff)

    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF3.List = T_Aff.Items
    End If

    Set T_Aff = Nothing
    Set Wksh = Nothing
End Sub

'This function shows the fourth list for the Geo (pretty much the same thing as done previously)
Sub ShowLst4(sAdm3 As String)

    [F_Geo].LST_Adm4.Clear

    Dim T_Aff As BetterArray
    Dim Wksh As Worksheet
    Dim sAdm1 As String
    Dim sAdm2 As String

    sAdm1 = [F_Geo].LST_Adm1.value
    sAdm2 = [F_Geo].LST_Adm2.value

    [F_Geo].TXT_Msg.value = [F_Geo].LST_Adm1.value & " | " & [F_Geo].LST_Adm2.value & " | " & [F_Geo].LST_Adm3.value

    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)
    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabAdm4), 1, sAdm1, 2, sAdm2, 3, sAdm3, returnIndex:=4)


    If T_Aff.Length > 0 Then
        [F_Geo].LST_Adm4.List = T_Aff.Items
    End If

    Set T_Aff = Nothing
    Set Wksh = Nothing
End Sub

'Fourth list of health facility
Sub ShowLstF4(sAdm3 As String)

    [F_Geo].LST_AdmF4.Clear

    Dim T_Aff As BetterArray
    Dim Wksh As Worksheet
    Dim sAdm1 As String
    Dim sAdm2 As String

    sAdm1 = [F_Geo].LST_AdmF1.value
    sAdm2 = [F_Geo].LST_AdmF2.value

    Set Wksh = ThisWorkbook.Worksheets(C_sSheetGeo)
    Set T_Aff = FilterLoTable(Wksh.ListObjects(C_sTabHF), 4, sAdm1, 3, sAdm2, 2, sAdm3, returnIndex:=1)

    [F_Geo].TXT_Msg.value = [F_Geo].LST_AdmF3.value & " | " & [F_Geo].LST_AdmF2.value & " | " & [F_Geo].LST_AdmF1.value

    If T_Aff.Length > 0 Then
        [F_Geo].LST_AdmF4.List = T_Aff.Items
    End If

    Set T_Aff = Nothing
    Set Wksh = Nothing
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
Sub SearchValue(ByVal sSearchedValue As String)
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
            [F_Geo].LST_ListeAgre.Clear
        End If
    Else
        If [F_Geo].LST_ListeAgre.ListCount - 1 <> T_Concat.UpperBound Then
            [F_Geo].LST_ListeAgre.List = T_Concat.Items
        End If
    End If

    Set T_result = Nothing
End Sub

Sub SeachHistoValue(sSearchedValue As String)
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
            [F_Geo].LST_Histo.Clear
        End If
    Else
        If [F_Geo].LST_Histo.ListCount - 1 <> T_HistoGeo.UpperBound Then
            [F_Geo].LST_Histo.List = T_HistoGeo.Items
        End If
    End If

    Set T_result = Nothing
End Sub

Sub SearchValueF(sSearchedValue As String)
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
            [F_Geo].LST_ListeAgreF.Clear
        End If
    Else
        If [F_Geo].LST_ListeAgreF.ListCount - 1 <> T_ConcatHF.UpperBound Then
            [F_Geo].LST_ListeAgreF.List = T_ConcatHF.Items
        End If
    End If

    Set T_result = Nothing
End Sub

Sub SeachHistoValueF(sSearchedValue As String)

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
            [F_Geo].LST_HistoF.Clear
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

Sub ClearOneHistoricGeobase(iGeoType As Byte)
    Dim WkshGeo As Worksheet
    Dim ShouldDelete As Integer

    Set WkshGeo = ThisWorkbook.Worksheets(C_sSheetGeo)

    ShouldDelete = MsgBox("Your historic geographic data  in the current workbook will be completely deleted, this action is irreversible. Proceed?", vbExclamation + vbYesNo, "Delete Historic")

    If ShouldDelete = vbYes Then
        If iGeoType = 0 Then
            If Not WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects(C_sTabHistoGeo).DataBodyRange.Delete
                [F_Geo].LST_Histo.Clear
            End If
        End If
        If iGeoType = 1 Then
            If Not WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
                WkshGeo.ListObjects(C_sTabHistoHF).DataBodyRange.Delete
                [F_Geo].LST_HistoF.Clear
            End If
        End If

        MsgBox "Done", vbinformation, "Delete Historic"
    End If
    'Add a message to say it is done
    Set WkshGeo = Nothing
End Sub
