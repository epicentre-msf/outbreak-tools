Attribute VB_Name = "LinelistCustomFunctions"
Option Explicit

'USER DEFINE FUNCTIONS FOR THE LINELIST ==========================================

'@Description("Get the date grande of column of type date which is "minimum date - maximum date")
'
'@Arguments:
'- Rng: a Range
'
'@Return:
'A String
'
'@Example: If "A1" contains date, DATE_RANGE("A1")

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
Public Function VALUE_OF(rng As Range, RngLook As Range, RngVal As Range) As String

    Dim sValLook As String
    Dim sSheetLook As String                     'Sheet name where to look for values
    Dim sSheetVal As String                      'Sheet name where to return the values

    Dim ColRngLook As Range                      'Column Range where to look for
    Dim ColRngVal  As Range                      'Column Range where to return the value from

    Dim iColLook As Long                         'columns to look and return values
    Dim sVal As String
    Dim iColVal As Long

    sValLook = rng.Value

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
            sVal = .Index(ColRngVal, .Match(sValLook, ColRngLook, 0))
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

'Epiweek function without specifying the year in select cases (works with all years)
Public Function Epiweek2(currentDate As Long) As Long
    Dim inDate As Long
    Dim firstDate As Long
    
    inDate = DateSerial(Year(currentDate), 1, 1)
    firstDate = inDate - Weekday(inDate, 2) + 1
     
    Epiweek2 = 1 + (currentDate - firstDate) \ 7

End Function

'Find the quarter, the year, the week or the month depending on the aggregation ============================================

'Quick function to define the aggregate
Private Function GetAgg(sAggregate As String) As String

    Select Case sAggregate

    Case TranslateLLMsg("MSG_Day")
        GetAgg = "day"
    Case TranslateLLMsg("MSG_Week")
        GetAgg = "week"
    Case TranslateLLMsg("MSG_Month")
        GetAgg = "month"
    Case TranslateLLMsg("MSG_Quarter")
        GetAgg = "quarter"
    Case TranslateLLMsg("MSG_Year")
        GetAgg = "year"
    Case Else                                    'Aggregate as week if unable to find the aggregate (defensive)
        GetAgg = "week"
    End Select

End Function

Public Function FindLastDay(sAggregate As String, inDate As Long) As Long

    Dim sAgg As String
    Dim dLastDay As Long
    Dim monthQuarter As Integer
    Dim monthDate As Integer

    sAgg = GetAgg(sAggregate)

    Select Case sAgg

    Case "day"

        dLastDay = inDate

    Case "week"

        dLastDay = inDate - Weekday(inDate, 2) + 7

    Case "month"

        dLastDay = DateSerial(Year(inDate), Month(inDate) + 1, 0)

    Case "quarter"

        monthDate = Month(inDate)
        monthQuarter = 3 * (IIf((monthDate Mod 3) = 0, ((monthDate - 1) \ 3), (monthDate \ 3))) + 1

        dLastDay = DateSerial(Year(inDate), monthQuarter + 3, 0)

    Case "year"

        dLastDay = DateSerial(Year(inDate) + 1, 1, 0)

    End Select

    FindLastDay = dLastDay

End Function

'Format a date to feet the aggregation selection ================================================================

Public Function FormatDateFromLastDay(sAggregate As String, inDate As Long, maxDate As Long, startDate As Long) As String

    Dim sAgg As String
    Dim sValue As String
    Dim monthDate As Integer
    Dim quarterDate As Integer

    sAgg = GetAgg(sAggregate)
    
    If startDate > maxDate Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If

    Select Case sAgg

    Case "day"
        sValue = Format(inDate, "dd-mmm-yyyy")
    Case "week"
        sValue = TranslateLLMsg("MSG_W") & IIf(Epiweek2(inDate) < 10, "0" & Epiweek2(inDate), Epiweek2(inDate)) & " - " & Year(inDate)
    Case "month"
        sValue = Format(inDate, "mmm - yyyy")
    Case "quarter"
        monthDate = Month(inDate)
        quarterDate = (IIf((monthDate Mod 3) = 0, ((monthDate - 1) \ 3), (monthDate \ 3))) + 1
        sValue = TranslateLLMsg("MSG_Q") & quarterDate & " - " & Year(inDate)
    Case "year"
        sValue = Year(inDate)
    End Select

    FormatDateFromLastDay = sValue
End Function

'Format a date range
Public Function FormatDateRange(MinDate As Long, maxDate As Long) As String

    FormatDateRange = Format(MinDate, "dd/mm/yyyy") & "-" & Format(maxDate, "dd/mm/yyyy")

End Function

