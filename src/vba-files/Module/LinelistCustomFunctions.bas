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

'Epiweek function without specifying the year in select cases (works with all years)
Public Function Epiweek2(currentDate As Long) As Long
    
    Dim inDate As Long
    Dim firstDate As Long
    Dim firstMondayDate As Long
    Dim borderLeftDate As Long
    Dim borderRightDate As Long
    Dim LastYearEpiWeek As Long

    inDate = DateSerial(Year(currentDate), 1, 1)
    firstMondayDate = inDate - Weekday(inDate, 2) + 1
    
    borderLeftDate = DateSerial(Year(currentDate) - 1, 12, 29)
    
    firstDate = IIf(firstMondayDate < borderLeftDate, firstMondayDate + 7, firstMondayDate)

    
    Epiweek2 = IIf(currentDate >= firstDate, 1 + (currentDate - firstDate) \ 7, Epiweek2(borderLeftDate))

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

Public Function FormatDateFromLastDay(sAggregate As String, startDate As Long, endDate As Long, MaxDate As Long) As String

    'enDate is the date of the end of the aggregation period
    'startDate is the startDate of the aggregation period
    'maxDate is the maximum Date of the time series

    Dim sAgg As String
    Dim sValue As String
    Dim monthDate As Integer
    Dim quarterDate As Integer

    sAgg = GetAgg(sAggregate)

    If startDate > MaxDate Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If

    Select Case sAgg

    Case "day"
        sValue = Format(endDate, "dd-mmm-yyyy")
    Case "week"
        sValue = TranslateLLMsg("MSG_W") & IIf(Epiweek2(endDate) < 10, "0" & Epiweek2(endDate), Epiweek2(endDate)) & " - " & Year(endDate)
    Case "month"
        sValue = Format(endDate, "mmm - yyyy")
    Case "quarter"
        monthDate = Month(endDate)
        quarterDate = (IIf((monthDate Mod 3) = 0, ((monthDate - 1) \ 3), (monthDate \ 3))) + 1
        sValue = TranslateLLMsg("MSG_Q") & quarterDate & " - " & Year(endDate)
    Case "year"
        sValue = Year(endDate)
    End Select

    FormatDateFromLastDay = sValue
End Function

'Format a date range
Public Function FormatDateRange(MinDate As Long, MaxDate As Long) As String

    FormatDateRange = Format(MinDate, "dd/mm/yyyy") & "-" & Format(MaxDate, "dd/mm/yyyy")

End Function

'Top admin levels and values for the tables on spatial analysis
Public Function TopAdminName(Admlevel As String, Admcount As Long) As String
    TopAdminName = vbNullString
End Function

Public Function TopAdminValue(admName As String, Admlevel As String, Admcount As Long) As Long
    TopAdminValue = 0
End Function

Public Function FirstAggDayFrom(endDate As Long, agg As String) As Long
    Dim firstDate As Long
    Dim timeAgg As String
    
    timeAgg = GetAgg(agg)
    
    Select Case timeAgg
    Case "day"
        firstDate = endDate - 53
    Case "week"
        firstDate = endDate - 371
    Case "month"
        firstDate = DateSerial(Year(endDate) - 4 - ((Month(endDate) - 5) \ 12), ((Month(endDate) - 5) Mod 12), 1) - 1
    Case "quarter"
        firstDate = DateSerial(Year(endDate) - 13 - ((Month(endDate) - 3) \ 12), ((Month(endDate) - 3) Mod 12), 1) - 1
    Case "year"
        firstDate = DateSerial(Year(endDate) - 53, Month(endDate), Day(endDate))
    End Select
    
    FirstAggDayFrom = firstDate
    
End Function

Public Function LastAggDayFrom(startDate As Long, agg As String) As Long
    Dim lastDate As Long
    Dim timeAgg As String
    
    timeAgg = GetAgg(agg)
    
    Select Case timeAgg
    Case "day"
        lastDate = startDate + 53
    Case "week"
        lastDate = startDate + 371
    Case "month"
        lastDate = DateSerial(Year(startDate) + 4 + ((Month(startDate) + 5) \ 12), ((Month(startDate) + 5) Mod 12) + 1, 1) - 1
    Case "quarter"
        lastDate = DateSerial(Year(startDate) + 13 + ((Month(startDate) + 3) \ 12), ((Month(startDate) + 3) Mod 12) + 1, 1) - 1
    Case "year"
        lastDate = DateSerial(Year(startDate) + 53, Month(startDate), Day(startDate))
    End Select
    
    LastAggDayFrom = lastDate
    
End Function

Public Function ValidMin(startDate As Long, endDate As Long, MinDate As Long, MaxDate As Long, agg As String) As Long
    
    Dim validation As Long
    Dim timeStamp As Long
    
    If startDate = 0 And endDate = 0 Then
        'Test if the minimum and the maximum are 0
        If MaxDate = 0 And MinDate = 0 Then
            validation = -1
        Else
            validation = MinDate
        End If
    ElseIf (startDate = 0) Then
        timeStamp = FirstAggDayFrom(endDate, agg)
        validation = Application.WorksheetFunction.Max(MinDate, timeStamp)
    Else
        validation = Application.WorksheetFunction.Max(MinDate, startDate)
    End If
    
    ValidMin = validation
End Function

Public Function ValidMax(startDate As Long, endDate As Long, MinDate As Long, MaxDate As Long, agg As String) As Long
    
    Dim validation As Long
    Dim timeStamp As Long
    
    'The two dates are equal to 0
    If startDate = 0 And endDate = 0 Then
        'Test if the minimum and the maximum are 0
        If MaxDate = 0 And MinDate = 0 Then
            validation = 1
        Else
            validation = MaxDate
        End If
    ElseIf (endDate = 0) Then
        timeStamp = LastAggDayFrom(startDate, agg)
        validation = Application.WorksheetFunction.Min(timeStamp, MaxDate)
    ElseIf (startDate = 0) Then
        validation = Application.WorksheetFunction.Min(endDate, MaxDate)
    ElseIf (startDate <> 0 And endDate <> 0) Then
        timeStamp = LastAggDayFrom(startDate, agg)
        validation = Application.WorksheetFunction.Min(timeStamp, endDate, MaxDate)
    End If
    
    ValidMax = validation
End Function


Public Function InfoUser(userDate As Long, actualDate As Long, Optional infotype As Byte = 1) As String
   
    Dim info As String
    If ((userDate <> actualDate) And (userDate <> 0)) Then
        info = IIf(infotype = 1, TranslateLLMsg("MSG_InfoStart"), TranslateLLMsg("MSG_InfoEnd"))
        InfoUser = info & " " & Format(actualDate, "dd/mm/yyyy")
    End If
    
End Function


