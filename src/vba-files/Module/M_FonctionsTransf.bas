Attribute VB_Name = "M_FonctionsTransf"
Option Explicit
'M_FonctionsTransf

Public Function IsEmptyTable(T_aTest) As Boolean

Dim test As Variant

IsEmptyTable = False
On Error GoTo crash
test = UBound(T_aTest)
On Error GoTo 0
Exit Function

crash:
IsEmptyTable = True

End Function

Sub QuickSort(T_aTrier, ByVal lngMin As Long, ByVal lngMax As Long)
 
    Dim strMidValue As String
    Dim lngHi As Long
    Dim lngLo As Long
    Dim lngIndex As Long
  
    If lngMin >= lngMax Then Exit Sub
  
    ' Valeur de partionnement
    lngIndex = Int((lngMax - lngMin + 1) * Rnd + lngMin)
    strMidValue = T_aTrier(lngIndex)
 
    ' Echanger les valeurs
    T_aTrier(lngIndex) = T_aTrier(lngMin)
 
    lngLo = lngMin
    lngHi = lngMax
    Do
        ' Chercher, à partir de lngHi, une valeur < strMidValue
        Do While T_aTrier(lngHi) >= strMidValue
            lngHi = lngHi - 1
            If lngHi <= lngLo Then Exit Do
        Loop
        If lngHi <= lngLo Then
            T_aTrier(lngLo) = strMidValue
            Exit Do
        End If
 
        ' Echanger les valeurs lngLo et lngHi
        T_aTrier(lngLo) = T_aTrier(lngHi)
 
        ' Chercher à partir de lngLo une valeur >= strMidValue
        lngLo = lngLo + 1
        Do While T_aTrier(lngLo) < strMidValue
            lngLo = lngLo + 1
            If lngLo >= lngHi Then Exit Do
        Loop
        If lngLo >= lngHi Then
            lngLo = lngHi
            T_aTrier(lngHi) = strMidValue
            Exit Do
        End If
 
        ' Echanger les valeurs lngLo et lngHi
        T_aTrier(lngHi) = T_aTrier(lngLo)
    Loop
 
    ' Trier les 2 sous-T_aTrieres
    QuickSort T_aTrier, lngMin, lngLo - 1
    QuickSort T_aTrier, lngLo + 1, lngMax
    
End Sub

Public Function LoadPathWindow() As String

Dim fDialog As Office.FileDialog

LoadPathWindow = ""
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
With fDialog
    .AllowMultiSelect = False
    .Title = "Chose your file"  'MSG_ChooseFile
    .Filters.Clear
    .Filters.Add "Feuille de calcul Excel", "*.xlsx, *.xlsm, *.xlsb,  *.xls"        'MSG_ExcelFile

    If .Show = True Then
        LoadPathWindow = .SelectedItems(1)
    End If
End With
Set fDialog = Nothing

End Function

Public Function LoadFolderWindow() As String

Dim fDialog As Office.FileDialog

LoadFolderWindow = ""
Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
With fDialog
    .AllowMultiSelect = False
    .Title = "Chose your directory"     'MSG_ChooseDir
    .Filters.Clear
    
    If .Show = True Then
        LoadFolderWindow = .SelectedItems(1)
    End If
End With
Set fDialog = Nothing

End Function

Public Function CleanSpecLettersInName(sName As String) As String     'supp tous les caract spéciaux du nom

Dim T_Caract
Dim i As Integer
Dim sRes As String

sRes = sName
T_Caract = [T_ascii]
i = 1
While i <= UBound(T_Caract, 1)
    sName = Replace(sName, T_Caract(i, 2), "")
    i = i + 1
Wend
CleanSpecLettersInName = sName

End Function

'                                                                       '
'_________________________ Liste des fonctions _________________________'

Public Function Epiweek(jour As Long) As Long
    
    Dim annee As Long
    
    Dim Jour0_2014, Jour0_2015, Jour0_2016, Jour0_2017, Jour0_2018, Jour0_2019, Jour0_2020, Jour0_2021, Jour0_2022 As Long

    Jour0_2014 = 41638
    Jour0_2015 = 42002
    Jour0_2016 = 42366
    Jour0_2017 = 42730
    Jour0_2018 = 43101
    Jour0_2019 = 43465
    Jour0_2020 = 43829
    Jour0_2021 = 44193
    Jour0_2022 = 44557
    annee = Year(jour)
    
    Select Case annee
     Case 2014
        Epiweek = 1 + Int((jour - Jour0_2014) / 7)
     Case 2015
        Epiweek = 1 + Int((jour - Jour0_2015) / 7)
     Case 2016
        Epiweek = 1 + Int((jour - Jour0_2016) / 7)
     Case 2017
        Epiweek = 1 + Int((jour - Jour0_2017) / 7)
     Case 2018
        Epiweek = 1 + Int((jour - Jour0_2018) / 7)
     Case 2019
        Epiweek = 1 + Int((jour - Jour0_2019) / 7)
     Case 2020
        Epiweek = 1 + Int((jour - Jour0_2020) / 7)
     Case 2021
        Epiweek = 1 + Int((jour - Jour0_2021) / 7)
    Case 2022
        Epiweek = 1 + Int((jour - Jour0_2022) / 7)
    End Select
    
End Function
