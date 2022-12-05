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

    Application.Volatile

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

        'There is only one table per worksheet, so I can just take the first listObject
        Set ColRngLook = ThisWorkbook.Worksheets(sSheetLook).ListObjects(1).ListColumns(iColLook).Range
        Set ColRngVal = ThisWorkbook.Worksheets(sSheetVal).ListObjects(1).ListColumns(iColVal).Range

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
'Quick function to define the aggregate
Private Function GetAgg(sAggregate As String) As String

    Dim rng As Range
    
    If ActiveSheet.Cells(1, 3).Value <> "TS-Analysis" Then
        GetAgg = "week"
        Exit Function
    End If
    
    Set rng = ActiveSheet.Range("TIME_UNIT_LIST")
    Select Case sAggregate

    Case rng.Cells(1, 1).Value
        GetAgg = "day"
    Case rng.Cells(2, 1).Value
        GetAgg = "week"
    Case rng.Cells(3, 1).Value
        GetAgg = "month"
    Case rng.Cells(4, 1).Value
        GetAgg = "quarter"
    Case rng.Cells(5, 1).Value
        GetAgg = "year"
    Case Else                                    'Aggregate as week if unable to find the aggregate (defensive)
        GetAgg = "week"
    End Select
 
End Function

Public Function FindLastDay(sAggregate As String, inDate As Long) As Long

    Application.Volatile
    
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

    Application.Volatile
    'enDate is the date of the end of the aggregation period
    'startDate is the startDate of the aggregation period
    'maxDate is the maximum Date of the time series

    Dim sAgg As String
    Dim sValue As String
    Dim monthDate As Integer
    Dim quarterDate As Integer
    Dim fun As WorksheetFunction
    
    
    
    If startDate > MaxDate Or (ActiveSheet.Cells(1, 3).Value <> "TS-Analysis") Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If
    

    sAgg = GetAgg(sAggregate)

    Set fun = Application.WorksheetFunction
    Select Case sAgg
    Case "day"
        sValue = Format(endDate, "dd-mmm-yyyy")
    Case "week"
        sValue = TranslateLLMsg("MSG_W") & fun.IsoWeekNum(endDate) & " - " & Year(endDate)
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

Public Function TopAdminValue(admname As String, Admlevel As String, Admcount As Long) As Long
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
    Application.Volatile
    
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
    Application.Volatile
    
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
    Application.Volatile
    
    Dim info As String
    If ((userDate <> actualDate) And (userDate <> 0)) Then
        info = IIf(infotype = 1, TranslateLLMsg("MSG_InfoStart"), TranslateLLMsg("MSG_InfoEnd"))
        InfoUser = info & " " & Format(actualDate, "dd/mm/yyyy")
    End If
    
End Function

'HF pcode
Public Function HFPCODE(cellRng As Range) As Variant
    Application.Volatile
    
    If ActiveSheet.Cells(1, 3).Value <> "HList" Then Exit Function
    
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    HFPCODE = geo.Pcode(LevelHF, cellRng)
End Function

'Geo Pcode
Public Function GEOPCODE(cellRng As Range, Level As Byte) As String
    Application.Volatile
    
    If ActiveSheet.Cells(1, 3).Value <> "HList" Then Exit Function
   
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    GEOPCODE = geo.Pcode(Level - 1, cellRng) '- 1 because 0 == admin1
    
End Function


Public Function GEOCONCAT(cellRng As Range, Level As Byte) As String
    Application.Volatile
    
    Dim concatValue As String

    Select Case Level
    
    Case 1

        concatValue = cellRng.Value

    Case 2

        concatValue = cellRng.Offset(, 1).Value & " | " & cellRng.Value

    Case 3

        concatValue = cellRng.Offset(, 2).Value & " | " & cellRng.Offset(, 1).Value & " | " & cellRng.Value

    Case 4

        concatValue = cellRng.Offset(, 3).Value & " | " & cellRng.Offset(, 2).Value & " | " & cellRng.Offset(, 1).Value & " | " & cellRng.Value

    Case Else

        concatValue = cellRng.Value

    End Select

    GEOCONCAT = concatValue
End Function

Public Function FindTopAdmin(adminLevel As String, adminOrder As Integer, varName As String) As String

    Dim geo As ILLGeo
    Dim sp As ILLSpatial
    Dim adminName As String
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets("Geo")
    Set geo = LLGeo.Create(sh)

    Set sh = ThisWorkbook.Worksheets("spatial_tables__")
    Set sp = LLSpatial.Create(sh)

    Select Case adminLevel

    Case geo.GeoNames("adm1_name")
    
        adminName = "adm1"
        
    Case geo.GeoNames("adm2_name")
    
        adminName = "adm2"
        
    Case geo.GeoNames("adm3_name")
    
        adminName = "adm3"
        
    Case geo.GeoNames("adm4_name")
    
        adminName = "adm4"
        
    Case Else
    
        adminName = "adm1"
        
    End Select

    FindTopAdmin = sp.FindTopValue(adminName, adminOrder, varName)

End Function

