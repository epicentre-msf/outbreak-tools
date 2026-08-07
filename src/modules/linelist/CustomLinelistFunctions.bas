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

'@section VALUE_OF lookup cache
'===============================================================================
'VALUE_OF runs once per formula cell on every recalculation pass, so it caches
'the two columns it reads instead of walking back to the sheet each time. Four
'slots, the least recently used one evicted, so several lookup tables stay warm
'at once. The state sits at module level rather than in Statics inside the
'function because ResetValueOfCache has to reach it.

Private Const VALUEOF_SLOT_COUNT As Long = 4

Private valueOfSheetName(1 To VALUEOF_SLOT_COUNT) As String
Private valueOfLookupIndex(1 To VALUEOF_SLOT_COUNT) As Long
Private valueOfValueIndex(1 To VALUEOF_SLOT_COUNT) As Long
Private valueOfKeys(1 To VALUEOF_SLOT_COUNT) As Variant
Private valueOfValues(1 To VALUEOF_SLOT_COUNT) As Variant
Private valueOfValid(1 To VALUEOF_SLOT_COUNT) As Boolean
Private valueOfLastUsed(1 To VALUEOF_SLOT_COUNT) As Long
Private valueOfCounter As Long


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


'@description Read the sheet type tag from worksheet-level HiddenNames.
'Lightweight helper similar to HiddenNameValue but reads from worksheet scope.
'Falls back to Cell(1,3) for legacy sheets that do not yet have the HiddenName.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, TS-Analysis, etc.).
Private Function SheetTag(ByVal sh As Worksheet) As String
    Dim raw As String

    SheetTag = vbNullString
    On Error Resume Next
    raw = sh.Names("sheet_type").RefersTo
    On Error GoTo 0

    If LenB(raw) > 0 Then
        If Left$(raw, 2) = "=" & Chr(34) Then
            SheetTag = Mid$(raw, 3, Len(raw) - 3)
        ElseIf Left$(raw, 1) = "=" Then
            SheetTag = Mid$(raw, 2)
        End If
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
'Reads the key column and the value column of the first ListObject on the named
'sheet once, parks them in one of four slots, and matches inside the cached
'arrays on every later call. ResetValueOfCache drops a slot when its sheet is
'edited, which is what keeps the cached copy honest.
'@param rng Range. Cell containing the lookup key.
'@param lookupSheetName String. Name of the worksheet hosting the lookup table.
'@param colLookupIndex Long. 1-based column index in the ListObject for the key.
'@param colValueIndex Long. 1-based column index in the ListObject for the result.
'@return Variant. Matched value, or vbNullString when not found.
'@EntryPoint
Public Function VALUE_OF(rng As Range, lookupSheetName As String, colLookupIndex As Long, colValueIndex As Long) As Variant
    Application.Volatile

    Dim lookupValue As Variant
    Dim matchPos As Variant
    Dim slot As Long

    ' Every path out of here answers an empty string, the error path included:
    ' a key cell holding #N/A must not turn the formula cell into #VALUE!.
    VALUE_OF = vbNullString
    On Error GoTo ErrorHandler

    lookupValue = rng.Value

    If lookupValue = vbNullString Then Exit Function
    If Trim$(lookupSheetName) = vbNullString Then Exit Function

    slot = CachedValueOfSlot(lookupSheetName, colLookupIndex, colValueIndex)
    If slot = 0 Then slot = LoadValueOfSlot(lookupSheetName, colLookupIndex, colValueIndex)
    If slot = 0 Then Exit Function

    ' Application.Match takes a VBA array as happily as it takes a Range, so
    ' the lookup keeps the case-insensitive matching the Range version had
    ' without going back to the sheet. The slot arrays are 1 To rowCount, so
    ' the position Match answers indexes the value array directly.
    matchPos = Application.Match(lookupValue, valueOfKeys(slot), 0)
    If IsError(matchPos) Then Exit Function

    VALUE_OF = valueOfValues(slot)(matchPos)
    Exit Function

ErrorHandler:
    VALUE_OF = vbNullString
End Function

'@description Drop the cached copy of a lookup table.
'Application.Volatile makes Excel call VALUE_OF again; it does not make the
'cache notice that the table underneath it moved. Without this the first read
'of a lookup table would answer every call for the rest of the session, and
'Ctrl+Alt+F9 would not clear it either, because the slots outlive a full
'recalculation. LinelistEventsManager.LLSheetChanged calls this with the name of
'the sheet that changed. Code that writes to a lookup sheet from inside another
'event handler, or with Application.EnableEvents off, gets no such event and
'has to call this itself.
'@param sheetName String. Name of the edited sheet. Blank drops every slot.
'@EntryPoint
Public Sub ResetValueOfCache(Optional ByVal sheetName As String = vbNullString)
    Dim slot As Long

    For slot = 1 To VALUEOF_SLOT_COUNT
        If LenB(sheetName) = 0 Then
            valueOfValid(slot) = False
        ElseIf StrComp(valueOfSheetName(slot), sheetName, vbTextCompare) = 0 Then
            valueOfValid(slot) = False
        End If
    Next slot
End Sub

'@description Find the slot already holding this sheet and column pair.
'A hit stamps the slot as the most recently used one, which is what the
'eviction walk reads.
'@param sheetName String. Worksheet hosting the lookup table.
'@param colLookupIndex Long. 1-based column index for the key.
'@param colValueIndex Long. 1-based column index for the result.
'@return Long. Slot number, or 0 when no slot matches.
Private Function CachedValueOfSlot(ByVal sheetName As String, _
                                   ByVal colLookupIndex As Long, _
                                   ByVal colValueIndex As Long) As Long
    Dim slot As Long

    For slot = 1 To VALUEOF_SLOT_COUNT
        If valueOfValid(slot) _
           And valueOfLookupIndex(slot) = colLookupIndex _
           And valueOfValueIndex(slot) = colValueIndex Then
            If StrComp(valueOfSheetName(slot), sheetName, vbTextCompare) = 0 Then
                valueOfCounter = valueOfCounter + 1
                valueOfLastUsed(slot) = valueOfCounter
                CachedValueOfSlot = slot
                Exit Function
            End If
        End If
    Next slot
End Function

'@description Pick the slot the next load may overwrite.
'The first free slot, or the least recently used one when all four are taken.
'@return Long. Slot number, always between 1 and VALUEOF_SLOT_COUNT.
Private Function ValueOfSlotToEvict() As Long
    Dim slot As Long
    Dim oldest As Long

    ValueOfSlotToEvict = 1
    oldest = valueOfLastUsed(1)

    For slot = 1 To VALUEOF_SLOT_COUNT
        If Not valueOfValid(slot) Then
            ValueOfSlotToEvict = slot
            Exit Function
        End If

        If valueOfLastUsed(slot) < oldest Then
            ValueOfSlotToEvict = slot
            oldest = valueOfLastUsed(slot)
        End If
    Next slot
End Function

'@description Read the two columns off the sheet into the evicted slot.
'Every way the sheet can disappoint answers 0, so VALUE_OF can treat a missing
'sheet, a sheet with no table and an out-of-range column index the same way.
'@param sheetName String. Worksheet hosting the lookup table.
'@param colLookupIndex Long. 1-based column index for the key.
'@param colValueIndex Long. 1-based column index for the result.
'@return Long. The slot now holding the data, or 0 when nothing was loaded.
Private Function LoadValueOfSlot(ByVal sheetName As String, _
                                 ByVal colLookupIndex As Long, _
                                 ByVal colValueIndex As Long) As Long
    Dim ws As Worksheet
    Dim Lo As ListObject
    Dim lookupRange As Range
    Dim valueRange As Range
    Dim keyList() As Variant
    Dim valueList() As Variant
    Dim block As Variant
    Dim rowCount As Long
    Dim r As Long
    Dim slot As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function
    If ws.ListObjects.Count = 0 Then Exit Function

    Set Lo = ws.ListObjects(1)

    If colLookupIndex < 1 Or colLookupIndex > Lo.ListColumns.Count Then Exit Function
    If colValueIndex < 1 Or colValueIndex > Lo.ListColumns.Count Then Exit Function

    Set lookupRange = Lo.ListColumns(colLookupIndex).DataBodyRange
    Set valueRange = Lo.ListColumns(colValueIndex).DataBodyRange

    If lookupRange Is Nothing Then Exit Function
    If valueRange Is Nothing Then Exit Function

    rowCount = lookupRange.Rows.Count
    ReDim keyList(1 To rowCount)
    ReDim valueList(1 To rowCount)

    If rowCount = 1 Then
        ' A one-row DataBodyRange answers .Value as a scalar rather than as a
        ' 2D array, and the walk below would raise a type mismatch on it.
        keyList(1) = lookupRange.Value
        valueList(1) = valueRange.Value
    Else
        block = lookupRange.Value
        For r = 1 To rowCount
            keyList(r) = block(r, 1)
        Next r

        block = valueRange.Value
        For r = 1 To rowCount
            valueList(r) = block(r, 1)
        Next r
    End If

    slot = ValueOfSlotToEvict
    valueOfCounter = valueOfCounter + 1

    valueOfSheetName(slot) = sheetName
    valueOfLookupIndex(slot) = colLookupIndex
    valueOfValueIndex(slot) = colValueIndex
    valueOfKeys(slot) = keyList
    valueOfValues(slot) = valueList
    valueOfValid(slot) = True
    valueOfLastUsed(slot) = valueOfCounter

    LoadValueOfSlot = slot
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
        If SheetTag(sh) = "HList" Then
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

    ' A blank cell reaches a Long parameter as 0, which is 1899-12-30, and the
    ' epi-year of that answers 1899. The formulas in the field guard the call
    ' with ISBLANK; the ones that do not used to get a date nobody typed.
    If currentDate <= 0 Then Exit Function

    ' Read week start from HiddenNames, default to Monday (1). The name is user
    ' data by the time a linelist is delivered, so a value that is not a number
    ' falls back rather than raising 13 out of a worksheet function as #VALUE!.
    rawStart = HiddenNameValue("RNG_EpiWeekStart", "1")
    If Not IsNumeric(rawStart) Then rawStart = "1"
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

    tagName = SheetTag(ActiveSheet)
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

    tagName = SheetTag(ActiveSheet)
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

Private Function EventService() As EventLinelist
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
