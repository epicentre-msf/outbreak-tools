Attribute VB_Name = "LinelistCustomFunctions"
Option Explicit

'@IgnoreModule IIfSideEffect

'USER DEFINE FUNCTIONS FOR THE LINELIST ========================================

Public Enum DayList
    Monday = 1
    TuesDay = 2
    Wednesday = 3
    Thursday = 4
    Friday = 5
    Saturday = 6
    Sunday = 0
End Enum


'@EntryPoint
Public Function DATE_RANGE(DateRng As Range) As String
    DATE_RANGE = Format(Application.WorksheetFunction.Min(DateRng), "DD/MM/YYYY") & _
                 " - " & Format(Application.WorksheetFunction.Max(DateRng), "DD/MM/YYYY")
End Function

'@EntryPoint
Public Function PLAGE_VALUE(rng1 As Range, rng2 As Range) As String
    PLAGE_VALUE = chr(13) & chr(10) & Format(rng1, "d-mmm-yyyy") & " " & ChrW(9472) & " " & Format(rng2, "d-mmm-yyyy")
End Function

'@EntryPoint
Public Function VALUE_OF(rng As Range, rngLook As Range, rngVal As Range) As Variant

    'Application.Volatile

    Dim sValLook As String
    Dim sSheetLook As String                     'Sheet name where to look for values
    Dim sSheetVal As String                      'Sheet name where to return the values

    Dim ColRngLook As Range                      'Column Range where to look for
    Dim ColRngVal  As Range                      'Column Range where to return the value from

    Dim iColLook As Long                         'columns to look and return values
    Dim retMatch As Variant
    Dim iColVal As Long
    Dim cellRngLook As Range
    Dim cellRngVal As Range

    sValLook = rng.Value


    If sValLook <> vbNullString Then
        sSheetLook = rngLook.Worksheet.Name
        sSheetVal = rngVal.Worksheet.Name

        iColLook = rngLook.Column
        iColVal = rngVal.Column

        'There is only one table per worksheet, so I can just take the first listObject
        Set ColRngLook = ThisWorkbook.Worksheets(sSheetLook).ListObjects(1).ListColumns(iColLook).Range
        Set ColRngVal = ThisWorkbook.Worksheets(sSheetVal).ListObjects(1).ListColumns(iColVal).Range
        Set cellRngLook = ColRngLook.Cells(1, 1)
        Set cellRngVal = ColRngVal.Cells(1, 1)

        Do While (cellRngLook.Value <> sValLook) And (cellRngLook.Row <= ColRngLook.Cells(ColRngLook.Rows.Count, 1).Row)
            Set cellRngLook = cellRngLook.Offset(1)
            Set cellRngVal = cellRngVal.Offset(1)
        Loop

        'Match
        If cellRngLook.Row <= ColRngLook.Cells(ColRngLook.Rows.Count, 1).Row Then retMatch = cellRngVal.Value

     End If

   VALUE_OF = retMatch
End Function

'@EntryPoint
Public Function ComputedOnFiltered() As String
    Application.Volatile
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim warningInfo As String
    Dim Lo As listObject
    Dim infoValue As String

    Set wb = ThisWorkbook
    warningInfo = wb.Worksheets("LinelistTranslation").Range("RNG_OnFiltered").Value

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            On Error Resume Next
                Set Lo = sh.ListObjects(1)
                If (Not Lo.AutoFilter Is Nothing) Then
                    infoValue = warningInfo
                    Exit For
                End If
            On Error GoTo 0
        End If
    Next

    ComputedOnFiltered = infoValue
End Function

'Epiweek function without specifying the year in select cases (works with all years)
'@EntryPoint
Public Function Epiweek(ByVal currentDate As Long, _ 
                        Optional ByVal weekStart As DayList = Monday) As Long

    Dim inDate As Long
    Dim firstDate As Long
    Dim firstDayDate As Long
    Dim borderLeftDate As Long

    inDate = DateSerial(Year(currentDate), 1, 1)
    firstDayDate = inDate - Weekday(inDate, weekStart + 1) + 1

    borderLeftDate = DateSerial(Year(currentDate) - 1, 12, 29)
    firstDate = IIf(firstDayDate < borderLeftDate, firstDayDate + 7, firstDayDate)

    If currentDate >= firstDate Then
        Epiweek = 1 + (currentDate - firstDate) \ 7
    Else
        Epiweek = Epiweek(borderLeftDate - 1)
    End If
End Function

'Find the quarter, the year, the week or the month depending on the aggregation

'Quick function to define the aggregate
Private Function GetAgg(sAggregate As String) As String

    Dim rng As Range
    Dim aggVal As String
    Dim tagName As String

    tagName = ActiveSheet.Cells(1, 3).Value
    If (tagName <> "TS-Analysis") And (tagName <> "SPT-Analysis") Then
        GetAgg = "week"
        Exit Function
    End If

    Set rng = Range("TIME_UNIT_LIST")
    Select Case sAggregate

    Case rng.Cells(1, 1).Value
        aggVal = "day"
    Case rng.Cells(2, 1).Value
        aggVal = "week"
    Case rng.Cells(3, 1).Value
        aggVal = "month"
    Case rng.Cells(4, 1).Value
        aggVal = "quarter"
    Case rng.Cells(5, 1).Value
        aggVal = "year"
    'Aggregate as week if unable to find the aggregate
    Case Else
        aggVal = "week"
    End Select

    GetAgg = aggVal
End Function

'@EntryPoint
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

'Format a date to feet the aggregation selection ===============================

'@EntryPoint
Public Function FormatDateFromLastDay(sAggregate As String, _ 
                                      startDate As Long, _ 
                                      endDate As Long, _ 
                                      MaxDate As Long) As String

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
    Dim quarterTag As String
    Dim lltradsh As Worksheet
    Dim weekTag As String
    Dim tagName As String

    tagName = ActiveSheet.Cells(1, 3).Value
    If startDate > MaxDate Or ((tagName <> "TS-Analysis") And (tagName <> "SPT-Analysis"))  Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If

    On Error Resume Next
        Set lltradsh = ThisWorkbook.Worksheets("LinelistTranslation")
        quarterTag = lltradsh.Range("RNG_Quarter").Value
        weekTag = lltradsh.Range("RNG_Week").Value
    On Error GoTo 0


    sAgg = GetAgg(sAggregate)

    Select Case sAgg
    Case "day"
        sValue = Format(endDate, "dd-mmm-yyyy")
    Case "week"
        epiW = Epiweek(endDate)
        epiYear = IIf(((epiW = 52 Or epiW = 53) And Month(endDate) = 1), _ 
                        Year(endDate) - 1, Year(endDate))
        sValue = weekTag & epiW & " - " & epiYear
    Case "month"
        sValue = Format(endDate, "mmm - yyyy")
    Case "quarter"
        monthDate = Month(endDate)
        quarterDate = (IIf((monthDate Mod 3) = 0, ((monthDate - 1) \ 3), _ 
                           (monthDate \ 3))) + 1
        sValue = quarterTag & quarterDate & " - " & Year(endDate)
    Case "year"
        sValue = Year(endDate)
    End Select

    FormatDateFromLastDay = sValue
End Function

'Format a date range
'@EntryPoint
Public Function FormatDateRange(MinDate As Long, MaxDate As Long) As String

    FormatDateRange = Format(MinDate, "dd/mm/yyyy") & "-" & _ 
                      Format(MaxDate, "dd/mm/yyyy")

End Function

'@EntryPoint
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

'@EntryPoint
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

'@EntryPoint
Public Function ValidMin(startDate As Long, endDate As Long, _ 
                         MinDate As Long, _ 
                         MaxDate As Long, agg As String) As Long
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

'@EntryPoint
Public Function ValidMax(startDate As Long, _ 
                         endDate As Long, _ 
                         MinDate As Long, MaxDate As Long, agg As String) As Long
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

'@EntryPoint
Public Function InfoUser(userDate As Long, _ 
                         actualDate As Long, _ 
                         Optional infotype As Byte = 1) As String
    Application.Volatile

    Dim lltradsh As Worksheet
    Dim infoStartTag As String
    Dim infoEndTag As String

    
    On Error Resume Next
        Set lltradsh = ThisWorkbook.Worksheets("LinelistTranslation")
        infoStartTag = lltradsh.Range("RNG_InfoStart").Value
        infoEndTag = lltradsh.Range("RNG_InfoEnd").Value
    On Error GoTo 0

    Dim info As String
    If ((userDate <> actualDate) And (userDate <> 0)) Then
        info = IIf(infotype = 1, infoStartTag, infoEndTag)
        InfoUser = info & " " & Format(actualDate, "dd/mm/yyyy")
    End If

End Function


'@EntryPoint
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

'@EntryPoint
Public Function FindTopAdmin(adminLevel As String, adminOrder As Integer, _ 
                             varName As String, _ 
                             Optional ByVal tabId As String = vbNullString) As String

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

'@EntryPoint
Public Function FindTopPop(adminLevel As String, adminOrder As Integer, _ 
                           varName As String, Optional ByVal tabId As String = vbNullString) As Long

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
'@EntryPoint
Public Function FindTopHF(adminOrder As Integer, varName As String, _ 
                         Optional ByVal tabId As String = vbNullString) As String

    Application.Volatile

    Dim sp As ILLSpatial
    Dim sh As Worksheet
    Dim actualVarName As String

    actualVarName = Split(varName, "_")(1)
    Set sh = ThisWorkbook.Worksheets("spatial_tables__")
    Set sp = LLSpatial.Create(sh)
    FindTopHF = sp.TopHFValue(adminOrder, actualVarName, tabId)
End Function
