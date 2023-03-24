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
Public Function PLAGE_VALUE(rng1 As Range, rng2 As Range) As String
    PLAGE_VALUE = chr(13) & chr(10) & Format(rng1, "d-mmm-yyyy") & " " & ChrW(9472) & " " & Format(rng2, "d-mmm-yyyy")
End Function

'
'
'
Public Function VALUE_OF(rng As Range, RngLook As Range, RngVal As Range) As Variant

    Application.Volatile

    Dim sValLook As String
    Dim sSheetLook As String                     'Sheet name where to look for values
    Dim sSheetVal As String                      'Sheet name where to return the values

    Dim ColRngLook As Range                      'Column Range where to look for
    Dim ColRngVal  As Range                      'Column Range where to return the value from

    Dim iColLook As Long                         'columns to look and return values
    Dim retMatch As Variant
    Dim iColVal As Long
    Dim indexMatch As Long

    sValLook = rng.Value


    If sValLook <> vbNullString Then
        sSheetLook = RngLook.Worksheet.Name
        sSheetVal = RngVal.Worksheet.Name

        iColLook = RngLook.Column
        iColVal = RngVal.Column

        'There is only one table per worksheet, so I can just take the first listObject
        Set ColRngLook = ThisWorkbook.Worksheets(sSheetLook).ListObjects(1).ListColumns(iColLook).Range
        Set ColRngVal = ThisWorkbook.Worksheets(sSheetVal).ListObjects(1).ListColumns(iColVal).Range
        indexMatch = -1

        On Error Resume Next
        With Application.WorksheetFunction
            indexMatch = .Match(sValLook, ColRngLook, 0)
            retMatch = .Index(ColRngVal, indexMatch)
        End With
        On Error GoTo 0
    End If

    If indexMatch <> -1 Then VALUE_OF = retMatch
End Function

'
'
Public Function ComputedOnFiltered() As String
    Application.Volatile
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim warningInfo As String
    Dim filteredSheet As String
    Dim Lo As ListObject
    Dim LoFiltered As ListObject
    Dim infoValue As String

    Set wb = ThisWorkbook
    warningInfo = wb.Worksheets("LinelistTranslation").Range("RNG_OnFiltered").Value

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
        filteredSheet = sh.Cells(1, 5).Value

        On Error Resume Next
            Set Lo = sh.ListObjects(1)
            Set LoFiltered = wb.Worksheets(filteredSheet).ListObjects(1)

            If Lo.Range.Rows.Count <> LoFiltered.Range.Rows.Count Then
                infoValue = warningInfo
                Exit For
            End If
        On Error GoTo 0
        End If
    Next

    ComputedOnFiltered = infoValue
End Function
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

    inDate = DateSerial(Year(currentDate), 1, 1)
    

    firstMondayDate = inDate - Weekday(inDate, 2) + 1
  
    borderLeftDate = DateSerial(Year(currentDate) - 1, 12, 29)
    
    firstDate = IIf(firstMondayDate < borderLeftDate, firstMondayDate + 7, firstMondayDate)

    If currentDate >= firstDate Then
        Epiweek2 = 1 + (currentDate - firstDate) \ 7
    Else
        Epiweek2 = Epiweek2(borderLeftDate - 1)
    End If

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
    Dim epiYear As Long
    Dim epiW As Long
    
    
    If startDate > MaxDate Or (ActiveSheet.Cells(1, 3).Value <> "TS-Analysis") Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If
    

    sAgg = GetAgg(sAggregate)

    Select Case sAgg
    Case "day"
        sValue = Format(endDate, "dd-mmm-yyyy")
    Case "week"
        epiW = Epiweek2(endDate)
        epiYear = IIf(((epiW = 52 Or epiW = 53) And Month(endDate) = 1), Year(endDate) - 1, Year(endDate))
        sValue = TranslateLLMsg("MSG_W") & epiW & " - " & epiYear
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

Public Function GeoPopulation(ByVal adminLevel As Byte, Optional ByVal concatValue As String = vbNullString) As Long
    Application.Volatile
    Dim geo As ILLGeo
    Dim popValue As String
    Dim popLng As Long
    
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    popLng = 0
    popValue = geo.Population(adminLevel, concatValue)

    On Error Resume Next
        popLng = CLng(popValue)
    On Error GoTo 0

    GeoPopulation = popLng
End Function
'There is No population for health facility

Public Function GEOCONCAT(cellRng As Range, Level As Byte) As String
    Application.Volatile
    
    Dim concatValue As String
    Dim nonEmptyValue As Boolean

    Select Case Level
    
    Case 1

        concatValue = cellRng.Value

    Case 2
        nonEmptyValue = (Not IsEmpty(cellRng)) And (Not IsEmpty(cellRng.Offset(, 1)))
        concatValue = IIf(nonEmptyValue, cellRng.Offset(, 1).Value & " | " & cellRng.Value, vbNullString)

    Case 3
        nonEmptyValue = (Not IsEmpty(cellRng)) And (Not IsEmpty(cellRng.Offset(, 1))) And (Not IsEmpty(cellRng.Offset(, 2)))
        concatValue = IIf(nonEmptyValue, cellRng.Offset(, 2).Value & " | " & cellRng.Offset(, 1).Value & " | " & cellRng.Value, vbNullString)

    Case 4

        nonEmptyValue = (Not IsEmpty(cellRng)) And (Not IsEmpty(cellRng.Offset(, 1))) And (Not IsEmpty(cellRng.Offset(, 2))) And (Not IsEmpty(cellRng.Offset(, 3)))
        concatValue = IIf(nonEmptyValue, cellRng.Offset(, 3).Value & " | " & cellRng.Offset(, 2).Value & " | " & cellRng.Offset(, 1).Value & " | " & cellRng.Value, vbNullString)

    Case Else

        concatValue = cellRng.Value

    End Select

    GEOCONCAT = concatValue
End Function

Public Function FindTopAdmin(adminLevel As String, adminOrder As Integer, varName As String, Optional ByVal tabId As String = vbNullString) As String

    Application.Volatile
    
    Dim geo As ILLGeo
    Dim sp As ILLSpatial
    Dim adminName As String
    Dim sh As Worksheet
    Dim actualVarName As String

    actualVarName = Split(varName, "_")(2)


    Set sh = ThisWorkbook.Worksheets("Geo")
    Set geo = LLGeo.Create(sh)

    Set sh = ThisWorkbook.Worksheets("spatial_tables__")
    Set sp = LLSpatial.Create(sh)
    adminName = geo.AdminCode(adminLevel)

    FindTopAdmin = sp.TopGeoValue(adminName, adminOrder, actualVarName, tabId)

End Function


Public Function FindTopPop(adminLevel As String, adminOrder As Integer, varName As String, Optional ByVal tabId As String = vbNullString) As Long

    Application.Volatile
    
    Dim geo As ILLGeo
    Dim sp As ILLSpatial
    Dim adminName As String
    Dim sh As Worksheet
    Dim actualVarName As String
    Dim pop As Long

    actualVarName = Split(varName, "_")(2)


    Set sh = ThisWorkbook.Worksheets("Geo")
    Set geo = LLGeo.Create(sh)

    Set sh = ThisWorkbook.Worksheets("spatial_tables__")
    Set sp = LLSpatial.Create(sh)
    adminName = geo.AdminCode(adminLevel)

    pop = 0
    On Error Resume Next
        pop = CLng(sp.TopGeoValue(adminName, adminOrder, actualVarName, tabId, 2))
    On Error GoTo 0

    FindTopPop = pop
End Function
    


'Find the corresponding value of a top admin for one variable

Public Function FindTopHF(adminOrder As Integer, varName As String, Optional ByVal tabId As String = vbNullString) As String

    Application.Volatile
    
    Dim sp As ILLSpatial
    Dim sh As Worksheet
    Dim actualVarName As String

    actualVarName = Split(varName, "_")(1)
    Set sh = ThisWorkbook.Worksheets("spatial_tables__")
    Set sp = LLSpatial.Create(sh)
    FindTopHF = sp.TopHFValue(adminOrder, actualVarName, tabId)
End Function


