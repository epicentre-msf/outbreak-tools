Attribute VB_Name = "LinelistCustomFunctions"


'USER DEFINE FUNCTIONS FOR THE LINELIST ==========================================

'@Description: Get the date grande of column of type date which is "minimum date - maximum date"
'
'@Arguments:
'- Rng: a Range
'
'@Return:
'A String
'
'@Example: If "A1" contains date, DATE_RANGE("A1")
'
Public Function DATE_RANGE(DateRng As Range) As String
        DATE_RANGE = Format(Application.WorksheetFunction.Min(DateRng), "DD/MM/YYYY") & _
        " - " & Format(Application.WorksheetFunction.Max(DateRng), "DD/MM/YYYY")
End Function


'
'
'
'
'
'
'
'
'
Public Function VALUE_OF(Rng As Range, RngLook As Range, RngVal As Range) As String

        Dim sValLook As String
        Dim sSheetLook As String 'Sheet name where to look for values
        Dim sSheetVal As String     'Sheet name where to return the values

        Dim ColRngLook As Range 'Column Range where to look for
        Dim ColRngVal  As Range 'Column Range where to return the value from

        Dim iColLook As Long 'columns to look and return values
        Dim sVal As String
        Dim iColVal As Long

        sValLook = Rng.value

        sVal = vbNullString

        If sValLook <> vbNullString Then
            sSheetLook = RngLook.Worksheet.Name
            sSheetVal = RngVal.Worksheet.Name

            iColLook = RngLook.Column
            iColVal = RngVal.Column

            Set ColRngLook = ThisWorkbook.Worksheets(sSheetLook).ListObjects(SheetListObjectName(sSheetLook)).ListColumns(iColLook).Range
            Set ColRngVal = ThisWorkbook.Worksheets(sSheetVal).ListObjects(SheetListObjectName(sSheetVal)).ListColumns(iColVal).Range

            On Error Resume Next
            With Application.WorksheetFunction
                sVal = .index(ColRngVal, .Match(sValLook, ColRngLook, 0))
            End With
            On Error GoTo 0
        End If

        VALUE_OF = sVal
End Function





'
'
'
'
'
'Epicemiological week function
Public Function Epiweek(jour As Long) As Long
    Dim annee As Long
    Dim Jour0_2014 As Long, Jour0_2015 As Long, Jour0_2016 As Long, Jour0_2017 As Long, Jour0_2018 As Long, Jour0_2019 As Long, Jour0_2020 As Long, Jour0_2021 As Long, Jour0_2022 As Long
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
