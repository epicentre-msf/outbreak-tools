Attribute VB_Name = "CustomLinelistFunctions"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect

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


'@section HiddenNames Helper
'===============================================================================

'@description Read a workbook-level HiddenName value as String.
'Lightweight helper that avoids creating a HiddenNames class instance,
'suitable for use in UDFs that are called many times during recalculation.
'@param nameId String. The HiddenName identifier (e.g. "RNG_EpiWeekStart").
'@param defaultValue String. Fallback when the name does not exist.
'@return String. The stored value, or defaultValue on failure.
Private Function HiddenNameValue(ByVal nameId As String, _
                                  ByVal defaultValue As String) As String
    Dim raw As String

    HiddenNameValue = defaultValue
    On Error Resume Next
    raw = ThisWorkbook.Names(nameId).RefersTo
    On Error GoTo 0

    If LenB(raw) = 0 Then Exit Function

    ' HiddenNames stores string values as ="value" and numeric values as =123
    If Left$(raw, 2) = "=" & Chr(34) Then
        HiddenNameValue = Mid$(raw, 3, Len(raw) - 3)
    ElseIf Left$(raw, 1) = "=" Then
        HiddenNameValue = Mid$(raw, 2)
    End If
End Function


'@section General UDFs
'===============================================================================

'@EntryPoint
Public Function DATE_RANGE(DateRng As Range) As String
    DATE_RANGE = Format(Application.WorksheetFunction.Min(DateRng), "DD/MM/YYYY") & _
                 " - " & Format(Application.WorksheetFunction.Max(DateRng), "DD/MM/YYYY")
End Function

'@EntryPoint
Public Function PLAGE_VALUE(rng1 As Range, rng2 As Range) As String
    PLAGE_VALUE = chr(13) & chr(10) & Format(rng1, "d-mmm-yyyy") & " " & ChrW(9472) & " " & Format(rng2, "d-mmm-yyyy")
End Function

'@description Lookup a value from a table on another sheet using cached arrays.
'Uses a 4-slot LRU cache to avoid re-reading sheet data on every recalculation.
'@param rng Range. Cell containing the lookup key.
'@param lookupSheetName String. Name of the worksheet hosting the lookup table.
'@param colLookupIndex Long. 1-based column index in the ListObject for the key.
'@param colValueIndex Long. 1-based column index in the ListObject for the result.
'@return Variant. Matched value, or vbNullString when not found.
'@EntryPoint
Public Function VALUE_OF(rng As Range, lookupSheetName As String, colLookupIndex As Long, colValueIndex As Long) As Variant
    Application.Volatile

    ' Cache slot 1
    Static slot1SheetName As String
    Static slot1LookupIndex As Long
    Static slot1ValueIndex As Long
    Static slot1LookupArray() As Variant
    Static slot1ValueArray() As Variant
    Static slot1Valid As Boolean
    Static slot1LastUsed As Long

    ' Cache slot 2
    Static slot2SheetName As String
    Static slot2LookupIndex As Long
    Static slot2ValueIndex As Long
    Static slot2LookupArray() As Variant
    Static slot2ValueArray() As Variant
    Static slot2Valid As Boolean
    Static slot2LastUsed As Long

    ' Cache slot 3
    Static slot3SheetName As String
    Static slot3LookupIndex As Long
    Static slot3ValueIndex As Long
    Static slot3LookupArray() As Variant
    Static slot3ValueArray() As Variant
    Static slot3Valid As Boolean
    Static slot3LastUsed As Long

    ' Cache slot 4
    Static slot4SheetName As String
    Static slot4LookupIndex As Long
    Static slot4ValueIndex As Long
    Static slot4LookupArray() As Variant
    Static slot4ValueArray() As Variant
    Static slot4Valid As Boolean
    Static slot4LastUsed As Long

    ' LRU tracking
    Static lruCounter As Long

    Dim ws As Worksheet
    Dim Lo As ListObject
    Dim lookupValue As Variant
    Dim i As Long
    Dim arraySize As Long
    Dim slotToUse As Long
    Dim rowCount As Long
    Dim r As Long
    Dim lruSlot As Long
    Dim lruMinValue As Long

    lookupValue = rng.Value

    If lookupValue = vbNullString Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    If Trim$(lookupSheetName) = vbNullString Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    ' Search for matching cache slot
    slotToUse = 0

    If slot1Valid And slot1SheetName = lookupSheetName And slot1LookupIndex = colLookupIndex And slot1ValueIndex = colValueIndex Then
        slotToUse = 1
        lruCounter = lruCounter + 1
        slot1LastUsed = lruCounter
    ElseIf slot2Valid And slot2SheetName = lookupSheetName And slot2LookupIndex = colLookupIndex And slot2ValueIndex = colValueIndex Then
        slotToUse = 2
        lruCounter = lruCounter + 1
        slot2LastUsed = lruCounter
    ElseIf slot3Valid And slot3SheetName = lookupSheetName And slot3LookupIndex = colLookupIndex And slot3ValueIndex = colValueIndex Then
        slotToUse = 3
        lruCounter = lruCounter + 1
        slot3LastUsed = lruCounter
    ElseIf slot4Valid And slot4SheetName = lookupSheetName And slot4LookupIndex = colLookupIndex And slot4ValueIndex = colValueIndex Then
        slotToUse = 4
        lruCounter = lruCounter + 1
        slot4LastUsed = lruCounter
    End If

    ' Cache miss - load data into LRU slot
    If slotToUse = 0 Then
        ' Find LRU slot (smallest lastUsed value, or first invalid slot)
        lruSlot = 1
        lruMinValue = slot1LastUsed
        If Not slot1Valid Then lruMinValue = -1

        If Not slot2Valid Then
            lruSlot = 2
            lruMinValue = -1
        ElseIf slot2LastUsed < lruMinValue Then
            lruSlot = 2
            lruMinValue = slot2LastUsed
        End If

        If Not slot3Valid Then
            lruSlot = 3
            lruMinValue = -1
        ElseIf slot3LastUsed < lruMinValue Then
            lruSlot = 3
            lruMinValue = slot3LastUsed
        End If

        If Not slot4Valid Then
            lruSlot = 4
        ElseIf slot4LastUsed < lruMinValue Then
            lruSlot = 4
        End If

        slotToUse = lruSlot

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(lookupSheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        If ws.ListObjects.Count = 0 Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        Set Lo = ws.ListObjects(1)

        If colLookupIndex < 1 Or colLookupIndex > Lo.ListColumns.Count Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        If colValueIndex < 1 Or colValueIndex > Lo.ListColumns.Count Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        If Lo.ListColumns(colLookupIndex).DataBodyRange Is Nothing Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        If Lo.ListColumns(colValueIndex).DataBodyRange Is Nothing Then
            VALUE_OF = vbNullString
            Exit Function
        End If

        ' Load columns into 2D arrays, then convert to 1D
        Dim lookupArray2D As Variant
        Dim valueArray2D As Variant

        lookupArray2D = Lo.ListColumns(colLookupIndex).DataBodyRange.Value
        valueArray2D = Lo.ListColumns(colValueIndex).DataBodyRange.Value
        rowCount = UBound(lookupArray2D, 1)

        lruCounter = lruCounter + 1

        Select Case slotToUse
            Case 1
                ReDim slot1LookupArray(1 To rowCount)
                ReDim slot1ValueArray(1 To rowCount)
                For r = 1 To rowCount
                    slot1LookupArray(r) = lookupArray2D(r, 1)
                    slot1ValueArray(r) = valueArray2D(r, 1)
                Next r
                slot1SheetName = lookupSheetName
                slot1LookupIndex = colLookupIndex
                slot1ValueIndex = colValueIndex
                slot1Valid = True
                slot1LastUsed = lruCounter

            Case 2
                ReDim slot2LookupArray(1 To rowCount)
                ReDim slot2ValueArray(1 To rowCount)
                For r = 1 To rowCount
                    slot2LookupArray(r) = lookupArray2D(r, 1)
                    slot2ValueArray(r) = valueArray2D(r, 1)
                Next r
                slot2SheetName = lookupSheetName
                slot2LookupIndex = colLookupIndex
                slot2ValueIndex = colValueIndex
                slot2Valid = True
                slot2LastUsed = lruCounter

            Case 3
                ReDim slot3LookupArray(1 To rowCount)
                ReDim slot3ValueArray(1 To rowCount)
                For r = 1 To rowCount
                    slot3LookupArray(r) = lookupArray2D(r, 1)
                    slot3ValueArray(r) = valueArray2D(r, 1)
                Next r
                slot3SheetName = lookupSheetName
                slot3LookupIndex = colLookupIndex
                slot3ValueIndex = colValueIndex
                slot3Valid = True
                slot3LastUsed = lruCounter

            Case 4
                ReDim slot4LookupArray(1 To rowCount)
                ReDim slot4ValueArray(1 To rowCount)
                For r = 1 To rowCount
                    slot4LookupArray(r) = lookupArray2D(r, 1)
                    slot4ValueArray(r) = valueArray2D(r, 1)
                Next r
                slot4SheetName = lookupSheetName
                slot4LookupIndex = colLookupIndex
                slot4ValueIndex = colValueIndex
                slot4Valid = True
                slot4LastUsed = lruCounter
        End Select
    End If

    ' Perform lookup in cached slot (1D arrays)
    Select Case slotToUse
        Case 1
            arraySize = UBound(slot1LookupArray)
            For i = 1 To arraySize
                If slot1LookupArray(i) = lookupValue Then
                    VALUE_OF = slot1ValueArray(i)
                    Exit Function
                End If
            Next i

        Case 2
            arraySize = UBound(slot2LookupArray)
            For i = 1 To arraySize
                If slot2LookupArray(i) = lookupValue Then
                    VALUE_OF = slot2ValueArray(i)
                    Exit Function
                End If
            Next i

        Case 3
            arraySize = UBound(slot3LookupArray)
            For i = 1 To arraySize
                If slot3LookupArray(i) = lookupValue Then
                    VALUE_OF = slot3ValueArray(i)
                    Exit Function
                End If
            Next i

        Case 4
            arraySize = UBound(slot4LookupArray)
            For i = 1 To arraySize
                If slot4LookupArray(i) = lookupValue Then
                    VALUE_OF = slot4ValueArray(i)
                    Exit Function
                End If
            Next i
    End Select

    ' No match found
    VALUE_OF = vbNullString
End Function

'@EntryPoint
Public Function ComputedOnFiltered() As String
    Application.Volatile
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim Lo As ListObject
    Dim filtCounter As Long

    Set wb = ThisWorkbook

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            On Error Resume Next
                Set Lo = sh.ListObjects(1)
                'Loop through all the filters in the listObject
                With Lo.AutoFilter.Filters
                    For filtCounter = 1 To .Count
                        If .Item(filtCounter).On Then GoTo AddWarning
                    Next
                End With
            On Error GoTo 0
        End If
    Next

    ComputedOnFiltered = vbNullString
    Exit Function

AddWarning:
    ComputedOnFiltered = HiddenNameValue("RNG_OnFiltered", vbNullString)
End Function


'@section Epidemiological Week
'===============================================================================

'@description Compute the start date of epidemiological week 1 for a given year.
'Week 1 is defined as the first week where 4 or more days fall in January,
'following the ISO 8601 convention generalised to any first day of the week.
'@param epiYear Long. The calendar year.
'@param weekStart Integer. First day of the week (DayList: 0=Sun, 1=Mon, ..., 6=Sat).
'@return Long. Date serial of the first day of week 1.
Private Function StartOfEpiWeek1(ByVal epiYear As Long, _
                                  ByVal weekStart As Integer) As Long
    Dim jan1 As Long
    Dim dayOfWeek As Long
    Dim weekStartDate As Long

    jan1 = DateSerial(epiYear, 1, 1)

    ' Weekday returns 1..7 where 1 = first day of the week
    ' weekStart+1 maps DayList values to VBA firstdayofweek parameter
    dayOfWeek = Weekday(jan1, weekStart + 1)

    ' Start of the week containing Jan 1
    weekStartDate = jan1 - dayOfWeek + 1

    ' If Jan 1 falls on day 5-7 of the week, fewer than 4 January days
    ' in this week, so week 1 starts the following week
    If dayOfWeek > 4 Then
        weekStartDate = weekStartDate + 7
    End If

    StartOfEpiWeek1 = weekStartDate
End Function

'@description Compute the formatted epidemiological week for a given date.
'Returns a string in the format W[week_number]-[year] where the week prefix
'is read from the HiddenName RNG_Week (language-dependent). The epi-year
'may differ from the calendar year at year boundaries. The first day of the
'week is read from the HiddenName RNG_EpiWeekStart (DayList values),
'defaulting to Monday (1) when unavailable.
'@param currentDate Long. Date serial number to compute the epiweek for.
'@param userStart Optional Integer. Override for the first day of the week.
'@return String. Formatted epiweek, e.g. "W1-2026".
'@EntryPoint
Public Function Epiweek(ByVal currentDate As Long, _
                         Optional ByVal userStart As Integer = -1) As String
    Application.Volatile

    Dim weekStart As Integer
    Dim weekTag As String
    Dim week1Start As Long
    Dim week1StartNext As Long
    Dim epiYear As Long
    Dim epiWeekNum As Long
    Dim rawStart As String

    ' Read week start from HiddenNames, default to Monday (1)
    rawStart = HiddenNameValue("RNG_EpiWeekStart", "1")
    weekStart = CInt(rawStart)

    ' Allow caller to override the week start
    If userStart >= 0 And userStart <= 6 Then
        weekStart = userStart
    End If

    ' Clamp to valid range (0=Sunday to 6=Saturday)
    If weekStart < 0 Or weekStart > 6 Then weekStart = 1

    ' Read translated week prefix (e.g. "W", "S" for Semaine, etc.)
    weekTag = HiddenNameValue("RNG_Week", "W")

    ' Compute start of week 1 for the current calendar year and the next year
    week1Start = StartOfEpiWeek1(Year(currentDate), weekStart)
    week1StartNext = StartOfEpiWeek1(Year(currentDate) + 1, weekStart)

    ' Determine which epi-year this date belongs to
    If currentDate >= week1StartNext Then
        ' Late December dates that belong to week 1 of the next year
        epiYear = Year(currentDate) + 1
        epiWeekNum = 1 + (currentDate - week1StartNext) \ 7
    ElseIf currentDate < week1Start Then
        ' Early January dates that belong to the last week of the previous year
        epiYear = Year(currentDate) - 1
        epiWeekNum = 1 + (currentDate - StartOfEpiWeek1(epiYear, weekStart)) \ 7
    Else
        epiYear = Year(currentDate)
        epiWeekNum = 1 + (currentDate - week1Start) \ 7
    End If

    Epiweek = weekTag & epiWeekNum & "-" & epiYear
End Function


'@section Aggregation Helpers
'===============================================================================

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

'@description Find the last day of an aggregation period containing inDate.
'Reads the epiweek start day from the workbook HiddenName RNG_EpiWeekStart.
'@param sAggregate String. The aggregation label (resolved via GetAgg).
'@param inDate Long. Date serial falling within the period.
'@return Long. Date serial of the last day of the aggregation period.
'@EntryPoint
Public Function FindLastDay(sAggregate As String, inDate As Long) As Long
    Application.Volatile

    Dim sAgg As String
    Dim dLastDay As Long
    Dim monthQuarter As Integer
    Dim monthDate As Integer
    Dim weekStart As Long
    Dim rawStart As String

    sAgg = GetAgg(sAggregate)

    ' Read week start from HiddenNames and add 1 for VBA Weekday parameter
    rawStart = HiddenNameValue("RNG_EpiWeekStart", "1")
    weekStart = CLng(rawStart) + 1

    Select Case sAgg

    Case "day"

        dLastDay = inDate

    Case "week"
        'replace the start of the week with the selected start for epiWeek
        dLastDay = inDate - Weekday(inDate, weekStart) + 7

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

'@description Format a date to match the aggregation selection.
'For weekly aggregation, delegates directly to Epiweek which returns
'the fully formatted string (e.g. "W1-2026").
'@param sAggregate String. The aggregation label.
'@param startDate Long. Start date of the aggregation period.
'@param endDate Long. End date of the aggregation period.
'@param MaxDate Long. Maximum date of the time series.
'@return String. Formatted date label for the aggregation period.
'@EntryPoint
Public Function FormatDateFromLastDay(sAggregate As String, _
                                      startDate As Long, _
                                      endDate As Long, _
                                      MaxDate As Long) As String

    Application.Volatile

    Dim sAgg As String
    Dim sValue As String
    Dim monthDate As Integer
    Dim quarterDate As Integer
    Dim quarterTag As String
    Dim tagName As String

    tagName = ActiveSheet.Cells(1, 3).Value
    If startDate > MaxDate Or ((tagName <> "TS-Analysis") And (tagName <> "SPT-Analysis"))  Then
        FormatDateFromLastDay = vbNullString
        Exit Function
    End If

    sAgg = GetAgg(sAggregate)

    Select Case sAgg
    Case "day"
        sValue = Format(endDate, "dd-mmm-yyyy")
    Case "week"
        ' Epiweek returns the fully formatted string (e.g. "W1-2026")
        sValue = Epiweek(endDate)
    Case "month"
        sValue = Format(endDate, "mmm - yyyy")
    Case "quarter"
        quarterTag = HiddenNameValue("RNG_Quarter", "Q")
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


'@section User Info
'===============================================================================

'@description Display a date info message when the user date differs from the actual date.
'Reads translated info tags from workbook-level HiddenNames (RNG_InfoStart, RNG_InfoEnd).
'@param userDate Long. Date entered by the user.
'@param actualDate Long. Computed actual date.
'@param infotype Optional Byte. 1 = start info, 2 = end info. Defaults to 1.
'@return String. Info message with formatted date, or vbNullString when dates match.
'@EntryPoint
Public Function InfoUser(userDate As Long, _
                         actualDate As Long, _
                         Optional infotype As Byte = 1) As String
    Application.Volatile

    Dim infoStartTag As String
    Dim infoEndTag As String
    Dim info As String

    infoStartTag = HiddenNameValue("RNG_InfoStart", vbNullString)
    infoEndTag = HiddenNameValue("RNG_InfoEnd", vbNullString)

    If ((userDate <> actualDate) And (userDate <> 0)) Then
        info = IIf(infotype = 1, infoStartTag, infoEndTag)
        InfoUser = info & " " & Format(actualDate, "dd/mm/yyyy")
    End If

End Function


'@section Geo / Spatial UDFs
'===============================================================================

Private Function EventService() As IEventLinelist
    Set EventService = LinelistEventsManager.EventLinelistService
End Function

'@EntryPoint
Public Function GEOCONCAT(cellRng As Range, Level As Byte) As String
    Application.Volatile
    GEOCONCAT = EventService.GeoConcat(cellRng, Level)
End Function

'@EntryPoint
Public Function FindTopAdmin(adminLevel As String, adminOrder As Integer, _
                             varName As String, _
                             Optional ByVal tabId As String = vbNullString) As String
    Application.Volatile
    FindTopAdmin = EventService.TopAdmin(adminLevel, adminOrder, varName, tabId)
End Function

'@EntryPoint
Public Function FindTopPop(adminLevel As String, adminOrder As Integer, _
                           varName As String, _
                           Optional ByVal tabId As String = vbNullString) As Long
    Application.Volatile
    FindTopPop = EventService.TopPop(adminLevel, adminOrder, varName, tabId)
End Function

'@EntryPoint
Public Function FindTopHF(adminOrder As Integer, varName As String, _
                          Optional ByVal tabId As String = vbNullString) As String
    Application.Volatile
    FindTopHF = EventService.TopHF(adminOrder, varName, tabId)
End Function
