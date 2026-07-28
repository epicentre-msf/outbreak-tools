VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLFormat"
Attribute VB_Description = "Formatting the linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Formatting the linelist")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument

Option Explicit
Private Const CLASSNAME As String = "LLFormat"

'@section Enumerations
'===============================================================================
'@description Supported formatting scopes. Each constant identifies a
'specific visual pattern that ApplyFormat applies to the supplied object.

Public Enum FormatScope
  AdminLargeButton = 1
  AllAnalysisSheet = 2
  AllLinelistSheet = 3
  AnalysisCategoriesBivariateColumn = 4
  AnalysisCategoriesBivariateRow = 5
  AnalysisCategoriesLabelColumn = 6
  AnalysisCategoriesLabelRow = 7
  AnalysisCategoriesUnivariateRow = 8
  AnalysisDateCell = 9
  AnalysisEndTable = 10
  AnalysisFirstRow = 11
  AnalysisGlobalSummaryRow = 12
  AnalysisInfoCell = 13
  AnalysisInfoCellLeft = 14
  AnalysisInteriorTable = 15
  AnalysisMissingColumn = 16
  AnalysisMissingRow = 17
  AnalysisMissingRowWithoutHeader = 18
  AnalysisOneCell = 19
  AnalysisOneCellCenter = 20
  AnalysisOneCellLeft = 21
  AnalysisPercent = 22
  AnalysisSection = 23
  AnalysisTableTitle = 24
  AnalysisTimeSeriesHeader = 25
  AnalysisUnivariateHeader = 26
  AnalysisWholeTable = 27
  CRFChoiceHeader = 28
  CRFChoiceValues = 29
  CRFDate = 30
  CRFDecimal = 31
  CRFFormatCustom = 32
  CRFInteger = 33
  CRFPercentage = 34
  CRFText = 35
  HListBorders = 36
  HListCalculatedFormulaCell = 37
  HListCalculatedFormulaHeader = 38
  HListCRFSection = 39
  HListCRFSheet = 40
  HListCRFSubSection = 41
  HListCustomPivotTables = 42
  HListGeo = 43
  HListGeoHeader = 44
  HListHeader = 45
  HListMainLab = 46
  HListNote = 47
  HListPrintedHeader = 48
  HListPrintedSheet = 49
  HListPrintedTable = 50
  HListSection = 51
  HListSimpleHeaderCell = 52
  HListSublab = 53
  HListSubSection = 54
  HListTable = 55
  LinelistGeoButton = 56
  LinelistGeoShape = 57
  LinelistHiddenCell = 58
  LinelistSelectionDropdown = 59
  LinelistShape = 60
  LinelistSmallButton = 61
  LinelistWarning = 62
  VListBorders = 63
  VListInnerBorders = 64
  VListMainLab = 65
  VListSection = 66
  VListSublab = 67
  VListSubSection = 68
End Enum

'@class LLFormat
'@description
'Reads design information from the LL format table and exposes helpers
'to apply those styles while caching lookups for performance. Supports
'import/export of format definitions across workbooks. Each FormatScope
'constant maps to a specific visual pattern applied to Worksheets,
'Ranges, Shapes, ListObjects, or Characters.
'@details
'The design values a format object reads do not move while the object is
'alive, so DesignValue caches the ANSWER and not only the row and column it
'came from. ApplyFormat is the hot path of a whole linelist build (127 call
'sites, 72 DesignValue calls in its own body), and without the answer cache
'every one of those calls crossed into Excel. A caller that edits the format
'sheet by hand and wants the new values asks for a new object, or goes
'through Import, which clears every cache.
'@depends CustomTable, Checking, BetterArray
'@version 1.0 (2026-02-09)


Private Const DEFAULT_DESIGN As String = "design 1"
Private Const COLUMN_LABEL As String = "Label"
Private Const DESIGN_TWO As String = "design 2"
Private Const DESIGN_USER_DEFINED As String = "user defined"
Private Const FALLBACK_COLOR As Long = vbBlack
Private Const FALLBACK_VALUE As Long = 0
Private Const CACHE_SEPARATOR As String = "|"

Private Type TLLFormat
    sheet As Worksheet
    design As String
    formatTable As ListObject
    labelColumnIndex As Long
    designColumnCache As Collection
    labelRowCache As Collection
    valueCache As Collection
    hasChecks As Boolean
    checks As Checking
    checkCounter As Long
    sealed As Boolean
End Type

Private this As TLLFormat

'@section Cache Management
'===============================================================================
'Lazy-init and invalidation of internal lookup caches.
'
'A keyed Collection is used, not a Dictionary: Scripting.Dictionary is not
'available on Mac Excel. UpdatedValues.RegistryNameCache uses the same shape.

'@label:reset-caches
'@sub-title Invalidate all cached references
Private Sub ResetCaches()
    Set this.formatTable = Nothing
    Set this.designColumnCache = Nothing
    Set this.labelRowCache = Nothing
    Set this.valueCache = Nothing
    'The label column index is a cache too, and it used to survive its own
    'invalidation because EnsureLabelColumn leaves early whenever it is above zero
    this.labelColumnIndex = 0
End Sub

'@label:ensure-format-ready
'@sub-title Ensure caches are initialised and the worksheet is set
Private Sub EnsureFormatReady()
    If this.sheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Format worksheet reference is missing"
    End If
    If this.designColumnCache Is Nothing Then Set this.designColumnCache = New Collection
    If this.labelRowCache Is Nothing Then Set this.labelRowCache = New Collection
    If this.valueCache Is Nothing Then Set this.valueCache = New Collection
End Sub

'@label:cached-value
'@sub-title Read one entry out of a keyed cache
'@param store Collection. The cache to read.
'@param cacheKey String. Key of the entry.
'@param found Boolean (ByRef). True when the key was present.
'@return Variant. The stored value, or Empty when the key is absent.
Private Function CachedValue(ByVal store As Collection, _
                             ByVal cacheKey As String, _
                             ByRef found As Boolean) As Variant
    Dim result As Variant

    found = False
    If store Is Nothing Then Exit Function

    'A Collection has no "does this key exist" test, so the miss is read off Err.
    'Err is cleared first: neither On Error Resume Next nor On Error GoTo 0
    'clears it, so a stale number from earlier code would read as a miss
    On Error Resume Next
        Err.Clear
        result = store.Item(cacheKey)
        found = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0

    If found Then CachedValue = result
End Function

'@label:store-cached-value
'@sub-title Add one entry to a keyed cache
'@param store Collection. The cache to write.
'@param cacheKey String. Key of the entry.
'@param cacheValue Variant. Value to keep.
Private Sub StoreCachedValue(ByVal store As Collection, _
                             ByVal cacheKey As String, _
                             ByVal cacheValue As Variant)
    If store Is Nothing Then Exit Sub

    'A duplicate key means another path already answered; keep the first answer
    On Error Resume Next
        store.Add cacheValue, cacheKey
    On Error GoTo 0
End Sub

'@section Format Table Access
'===============================================================================
'One owner for each rule about what a valid format sheet looks like.

'@label:table-of-sheet
'@sub-title Retrieve the first ListObject of a worksheet
'@param sh Worksheet. The worksheet to read.
'@return ListObject. The first table on the sheet.
Private Function TableOfSheet(ByVal sh As Worksheet) As ListObject
    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Format worksheet must be supplied"
    End If

    If sh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "No format table found on worksheet '" & sh.Name & "'"
    End If

    Set TableOfSheet = sh.ListObjects(1)
End Function

'@label:header-of-table
'@sub-title Retrieve the header row of a format table
'@param formatLo ListObject. The format table.
'@return Range. The header row.
Private Function HeaderOfTable(ByVal formatLo As ListObject) As Range
    Dim header As Range

    On Error Resume Next
        Set header = formatLo.HeaderRowRange
    On Error GoTo 0

    If header Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Format table on worksheet '" & formatLo.Parent.Name & "' is missing headers"
    End If

    Set HeaderOfTable = header
End Function

'@label:ensure-format-sheet-ready
'@sub-title Validate a worksheet as a valid format source
'@param sh Worksheet. The worksheet to validate.
'@param resolvedDesign String (ByRef). Falls back to DEFAULT_DESIGN when the requested design column is absent.
Private Sub EnsureFormatSheetReady(ByVal sh As Worksheet, ByRef resolvedDesign As String)

    Dim formatLo As ListObject
    Dim header As Range

    Set formatLo = TableOfSheet(sh)
    Set header = HeaderOfTable(formatLo)

    'Two tests, not one joined by Or: VBA runs both sides of Or, so the old
    'single line read Rows.Count off Nothing and raised 91 instead of the message
    If formatLo.DataBodyRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Format table '" & formatLo.Name & "' does not contain any data rows"
    End If

    If formatLo.DataBodyRange.Rows.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Format table '" & formatLo.Name & "' does not contain any data rows"
    End If

    If Not ColumnExists(header, COLUMN_LABEL) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Column '" & COLUMN_LABEL & "' is required in format table '" & formatLo.Name & "'"
    End If

    If Not ColumnExists(header, DEFAULT_DESIGN) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Default design '" & DEFAULT_DESIGN & "' was not found in format table '" & formatLo.Name & "'"
    End If

    If Not ColumnExists(header, resolvedDesign) Then
        resolvedDesign = DEFAULT_DESIGN
    End If
End Sub

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking object
'@param label String. The message to record.
Private Sub LogInfo(ByVal label As String)
    If this.checks Is Nothing Then
        Set this.checks = Checking.Create("LLFormat trace")
    End If
    this.checks.Add CStr(this.checkCounter), label, checkingNote
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:format-table
'@sub-title Retrieve the first ListObject on the format worksheet
'@return ListObject. The format table.
Private Function FormatTable() As ListObject
    EnsureFormatReady
    If this.formatTable Is Nothing Then
        Set this.formatTable = TableOfSheet(this.sheet)
    End If
    Set FormatTable = this.formatTable
End Function

'@label:header-range
'@sub-title Retrieve the header row range of the format table
'@return Range. The header row.
Private Function HeaderRange() As Range
    Set HeaderRange = HeaderOfTable(FormatTable())
End Function

'@label:ensure-label-column
'@sub-title Cache the column index of the Label column
Private Sub EnsureLabelColumn()
    If this.labelColumnIndex > 0 Then Exit Sub
    this.labelColumnIndex = ResolveColumnIndex(HeaderRange(), COLUMN_LABEL)
End Sub

'@label:find-column-index
'@sub-title Find the column index of a named header, or zero when it is absent
'@param header Range. The header row to search.
'@param columnName String. Column caption to locate.
'@return Long. The 1-based column index, or 0 when not found.
Private Function FindColumnIndex(ByVal header As Range, ByVal columnName As String) As Long
    Dim matchCell As Range

    On Error Resume Next
        Set matchCell = header.Find(What:=columnName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If Not matchCell Is Nothing Then FindColumnIndex = matchCell.Column
End Function

'@label:column-exists
'@sub-title Test whether a named column exists in a header range
'@param header Range. The header row to search.
'@param columnName String. Column caption to test.
'@return Boolean. True when the column is found.
Private Function ColumnExists(ByVal header As Range, ByVal columnName As String) As Boolean
    ColumnExists = (FindColumnIndex(header, columnName) > 0)
End Function

'@label:resolve-column-index
'@sub-title Find the column index of a named header, raising when it is absent
'@param header Range. The header row to search.
'@param columnName String. Column caption to locate.
'@return Long. The 1-based column index.
Private Function ResolveColumnIndex(ByVal header As Range, ByVal columnName As String) As Long
    ResolveColumnIndex = FindColumnIndex(header, columnName)
    If ResolveColumnIndex = 0 Then
        ThrowError ProjectError.ElementNotFound, "Column '" & columnName & "' missing from format table"
    End If
End Function

'@label:design-column-index
'@sub-title Resolve and cache the column index for a design name
'@param designName String. Name of the design column.
'@return Long. The 1-based column index.
Private Function DesignColumnIndex(ByVal designName As String) As Long
    Dim cached As Variant
    Dim found As Boolean

    EnsureFormatReady

    cached = CachedValue(this.designColumnCache, LCase$(designName), found)
    If found Then
        DesignColumnIndex = CLng(cached)
        Exit Function
    End If

    DesignColumnIndex = ResolveColumnIndex(HeaderRange(), designName)
    StoreCachedValue this.designColumnCache, LCase$(designName), DesignColumnIndex
End Function

'@label:label-row-index
'@sub-title Resolve and cache the row index for a format label
'@details
'The lookup tries an exact match first and only then a partial one. Four of
'the labels this class asks for are stored in the format table with extra
'words around them - "gap for graphs" sits in "default row height gap for
'graphs", "dropdown interior color" in "Select dropdown interior color",
'"time series header font" in "Time series header font color" and "time
'series header interior" in "Time series header interior color" - so partial
'matching cannot simply be dropped. Trying the whole string first means a new
'row whose label merely CONTAINS an existing one can no longer take it over.
'@param labelName String. Label text to locate in the Label column.
'@return Long. The 1-based row index, or 0 when not found.
Private Function LabelRowIndex(ByVal labelName As String) As Long
    Dim cached As Variant
    Dim found As Boolean
    Dim dataColumn As Range
    Dim matchCell As Range

    EnsureFormatReady
    EnsureLabelColumn

    cached = CachedValue(this.labelRowCache, LCase$(labelName), found)
    If found Then
        LabelRowIndex = CLng(cached)
        Exit Function
    End If

    Set dataColumn = FormatTable().ListColumns(COLUMN_LABEL).DataBodyRange

    On Error Resume Next
        Set matchCell = dataColumn.Find(What:=labelName, LookAt:=xlWhole, MatchCase:=False)
        If matchCell Is Nothing Then
            Set matchCell = dataColumn.Find(What:=labelName, LookAt:=xlPart, MatchCase:=False)
        End If
    On Error GoTo 0

    If Not matchCell Is Nothing Then LabelRowIndex = matchCell.Row

    'The miss is cached and logged here, exactly once per label. A label that is
    'never found used to re-scan the whole Label column and add one more log line
    'on every single call, and a typo inside a build loop meant thousands of both
    If LabelRowIndex = 0 Then
        LogInfo "Label '" & labelName & "' not found; returning fallback value"
    End If

    StoreCachedValue this.labelRowCache, LCase$(labelName), LabelRowIndex
End Function

'@section Factory
'===============================================================================
'PredeclaredId factory method.

'@label:create
'@sub-title Instantiate an LL format engine for a worksheet
'@details
'Validates the format worksheet, resolves the design column, and
'returns a fully initialised LLFormat instance.
'@param formatSh Worksheet. Worksheet containing the format specification table.
'@param designType String. Design column to use. Defaults to DEFAULT_DESIGN.
'@return LLFormat. A fully initialised LLFormat instance.
'@export
Public Function Create(ByVal formatSh As Worksheet, _
                       Optional ByVal designType As String = DEFAULT_DESIGN) As LLFormat

    Dim resolvedDesign As String

    resolvedDesign = designType

    EnsureFormatSheetReady formatSh, resolvedDesign

    Dim newFormat As LLFormat
    Set newFormat = New LLFormat

    Set newFormat.Wksh = formatSh
    newFormat.DesignName = resolvedDesign
    newFormat.Seal
    Set Create = newFormat

End Function

'@section Public Interface
'===============================================================================
'Public properties exposed by the factory and the interface.

'@label:wksh
'@prop-title Retrieve the worksheet hosting the format table
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sheet
End Property

'@label:wksh-set
'@prop-title Assign the worksheet backing this formatter
'@details Clears all caches when the worksheet changes.
'@param sh Worksheet. The new host worksheet.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sheet = sh
    ResetCaches
End Property

'@label:design-name
'@prop-title Name of the active design column
'@return String. The active design name.
Public Property Get DesignName() As String
    DesignName = this.design
End Property

'@label:design-name-set
'@prop-title Change the active design column
'@details Resets cached lookups when the design changes.
'@param design String. The new design column name.
Public Property Let DesignName(ByVal design As String)
    GuardNotSealed "DesignName"
    this.design = design
    ResetCaches
End Property

'@label:has-checkings
'@prop-title Whether the format object has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Checking. The checking instance, or Nothing.
Public Property Get CheckingValues() As Checking
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section Design Lookups
'===============================================================================
'Retrieve raw style values from the format table.

'@label:design-value
'@sub-title Look up a design value or colour from the format table
'@details
'Retrieves either the fill colour (Interior.Color) or the raw cell
'value for the specified label row under the active design column.
'Returns a fallback when the label is not found and logs a diagnostic.
'The answer is cached for the life of the object; see the class @details.
'@param labelName String. Label text stored in the format table.
'@param returnColor Boolean. When True returns Interior.Color; otherwise the raw value. Defaults to True.
'@return Variant. The colour (Long) or cell value.
'@export
Public Function DesignValue(ByVal labelName As String, _
                             Optional ByVal returnColor As Boolean = True) As Variant

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetCell As Range
    Dim resultValue As Variant
    Dim cacheKey As String
    Dim cached As Variant
    Dim found As Boolean

    EnsureFormatReady

    'The key carries the design and returnColor as well as the label. The two
    'fallbacks are different constants, and the same cell answers with its fill
    'colour or with its value depending on returnColor, so the label alone would
    'hand one caller the other caller's answer
    cacheKey = LCase$(this.design) & CACHE_SEPARATOR & CStr(returnColor) & _
               CACHE_SEPARATOR & LCase$(labelName)

    cached = CachedValue(this.valueCache, cacheKey, found)
    If found Then
        DesignValue = cached
        Exit Function
    End If

    rowIndex = LabelRowIndex(labelName)
    If rowIndex = 0 Then
        DesignValue = IIf(returnColor, FALLBACK_COLOR, FALLBACK_VALUE)
        StoreCachedValue this.valueCache, cacheKey, DesignValue
        Exit Function
    End If

    columnIndex = DesignColumnIndex(DesignName())
    Set targetCell = Wksh.Cells(rowIndex, columnIndex)

    If returnColor Then
        resultValue = targetCell.Interior.Color
        If CLng(resultValue) = 0 Then resultValue = FALLBACK_COLOR
    Else
        resultValue = targetCell.Value
        If IsEmpty(resultValue) Then resultValue = FALLBACK_VALUE
    End If

    DesignValue = resultValue
    StoreCachedValue this.valueCache, cacheKey, resultValue
End Function

'@label:analysis-base-font-size
'@sub-title Base font size used by the analysis format scopes
'@return Variant. The design font size, or 11 when the design cell is blank.
Private Function AnalysisBaseFontSize() As Variant
    Dim styleValue As Variant

    styleValue = DesignValue("analysis base font size", False)
    AnalysisBaseFontSize = IIf(styleValue = 0, 11, styleValue)
End Function

'@section Drawing Helpers
'===============================================================================
'Borders, fills and fonts.

'@label:draw-lines
'@sub-title Draw border lines around a range
'@param rng Range. Target range.
'@param At String. Edge selector (All, AvoidLeft, AvoidRight, etc.). Defaults to "All".
'@param weight Integer. Line weight constant. Defaults to xlHairline.
'@param line Integer. Line style constant. Defaults to xlContinuous.
'@param color Variant. Border colour. Defaults to vbBlack.
Private Sub DrawLines(ByVal rng As Range, _
                      Optional At As String = "All", _
                      Optional weight As Integer = xlHairline, _
                      Optional line As Integer = xlContinuous, _
                      Optional color As Variant = vbBlack)

    'A fixed local array, not a BetterArray. This sub is written 50 times inside
    'the format branches and runs far more often than that, so building an object
    'to hold between one and six constants was the whole cost of the call
    Dim borderPositions(1 To 6) As Long
    Dim borderCount As Long
    Dim counter As Long

    Select Case At

    Case "All"
        borderCount = 6
        borderPositions(1) = xlEdgeLeft
        borderPositions(2) = xlEdgeRight
        borderPositions(3) = xlEdgeBottom
        borderPositions(4) = xlEdgeTop
        borderPositions(5) = xlInsideHorizontal
        borderPositions(6) = xlInsideVertical
    Case "AvoidLeft", "AvoidRight"
        borderCount = 3
        borderPositions(1) = xlEdgeBottom
        borderPositions(2) = xlEdgeTop
        borderPositions(3) = xlInsideHorizontal
    Case "AvoidTop", "AvoidBottom"
        borderCount = 3
        borderPositions(1) = xlEdgeLeft
        borderPositions(2) = xlEdgeRight
        borderPositions(3) = xlInsideVertical
    Case "Left"
        borderCount = 1
        borderPositions(1) = xlEdgeLeft
    Case "Right"
        borderCount = 1
        borderPositions(1) = xlEdgeRight
    Case "Bottom"
        borderCount = 1
        borderPositions(1) = xlEdgeBottom
    Case "Top"
        borderCount = 1
        borderPositions(1) = xlEdgeTop
    Case Else
        borderCount = 1
        borderPositions(1) = xlEdgeBottom
    End Select

    With rng
        For counter = 1 To borderCount
            With .Borders(borderPositions(counter))
                .weight = weight
                .LineStyle = line
                .color = color
                .TintAndShade = 0.2
            End With
        Next counter
    End With
End Sub

'@label:draw-border
'@sub-title Draw a rectangular border around a range
'@details
'Writes the four properties in the SAME order as DrawLines. Several scopes
'used to call DrawLines four times, once per edge, and now call this sub once
'instead; with line:=xlDouble the LineStyle write is what settles the final
'weight, so a different order would make those frames come out heavier.
'@param rng Range. Target range to border.
'@param weight Optional Integer. Line weight constant. Defaults to xlThin.
'@param color Optional Variant. Border colour. Defaults to vbBlack.
'@param line Optional Long. Line style constant. Defaults to xlContinuous.
Private Sub DrawBorder(ByVal rng As Range, Optional weight As Integer = xlThin, _
                       Optional color As Variant = vbBlack, _
                       Optional line As Long = xlContinuous)

    Dim counter As Integer
    For counter = 7 To 10                        'xltop, left, right and bottom
        With rng.Borders(counter)
            .weight = weight
            .LineStyle = line
            .color = color
            .TintAndShade = 0.2
        End With
    Next counter
End Sub

'@label:format-range
'@sub-title Apply fill, font, alignment, and number format to a range
'@details
'The Variant arguments are compared against vbNullString to mean "leave this
'property alone". That works because VBA sorts every number below every
'string, and several scopes pass vbNullString on purpose to skip a property.
'@param rng Range. Target range.
'@param fillValue Optional String. Cell value to write. Defaults to "".
'@param interiorColor Optional Variant. Interior fill colour. Defaults to "".
'@param fontColor Optional Variant. Font colour. Defaults to "".
'@param isBold Optional Boolean. Bold font. Defaults to False.
'@param Horiz Optional Integer. Horizontal alignment. Defaults to xlHAlignCenter.
'@param Verti Optional Integer. Vertical alignment. Defaults to xlVAlignCenter.
'@param FontSize Optional Double. Font size in points. Defaults to 0 (unchanged).
'@param NumFormat Optional String. Number format string. Defaults to "".
'@param wrap Optional String. When "wrap", enables text wrapping. Defaults to vbNullString.
Private Sub FormatRange(ByVal rng As Range, _
                        Optional ByVal fillValue As String = "", _
                        Optional ByVal interiorColor As Variant = "", _
                        Optional ByVal fontColor As Variant = "", _
                        Optional ByVal isBold As Boolean = False, _
                        Optional ByVal Horiz As Integer = xlHAlignCenter, _
                        Optional ByVal Verti As Integer = xlVAlignCenter, _
                        Optional ByVal FontSize As Double = 0, _
                        Optional ByVal NumFormat As String = "", _
                        Optional ByVal wrap As String = vbNullString)

    With rng
        If interiorColor <> vbNullString Then .Interior.color = interiorColor
        If fontColor <> vbNullString Then .Font.color = fontColor
        If isBold Then .Font.Bold = isBold
        If FontSize <> 0 Then .Font.Size = FontSize
        If NumFormat <> vbNullString Then .NumberFormat = NumFormat
        .HorizontalAlignment = Horiz
        .VerticalAlignment = Verti
        If fillValue <> vbNullString Then .Value = fillValue
        If wrap = "wrap" Then .WrapText = True
    End With
End Sub

'@label:format-characters
'@sub-title Apply font styling to a Characters object
'@param Char Characters. The Characters object to style.
'@param fontColor Optional Variant. Font colour. Defaults to "".
'@param isBold Optional Boolean. Bold font. Defaults to False.
'@param FontSize Optional Double. Font size in points. Defaults to 0 (unchanged).
Private Sub FormatCharacters(ByVal Char As Characters, _
                             Optional ByVal fontColor As Variant = "", _
                             Optional ByVal isBold As Boolean = False, _
                             Optional ByVal FontSize As Double = 0)

    With Char
        If fontColor <> vbNullString Then .Font.color = fontColor
        .Font.Bold = isBold
        If FontSize <> 0 Then .Font.Size = FontSize
    End With
End Sub

'@label:format-shape
'@sub-title Apply fill and font styling to a Shape object
'@param shp Shape. The Shape to style.
'@param fontColor Optional Variant. Text colour. Defaults to vbNullString.
'@param interiorColor Optional Variant. Fill colour. Defaults to vbNullString.
'@param isBold Optional Boolean. Bold font. Defaults to False.
'@param FontSize Optional Double. Font size in points. Defaults to 0 (unchanged).
Private Sub FormatShape(ByVal shp As Shape, _
                        Optional ByVal fontColor As Variant = vbNullString, _
                        Optional ByVal interiorColor As Variant = vbNullString, _
                        Optional ByVal isBold As Boolean = False, _
                        Optional ByVal FontSize As Double = 0)
    With shp
        'Fill color
        If interiorColor <> vbNullString Then
            .Fill.ForeColor.RGB = interiorColor
            .Fill.BackColor.RGB = interiorColor
        End If

        With .TextFrame2
            'Text alignment
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoTrue

            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter

                'Font Zize and Font Shape
                If FontSize <> 0 Then .Font.Size = FontSize
                If isBold Then .Font.Bold = isBold
                'Text Color
                If fontColor <> vbNullString Then .Font.Fill.ForeColor.RGB = fontColor
            End With
        End With
    End With
End Sub

'@label:remove-gridlines
'@sub-title Hide gridlines on a worksheet
'@param sh Worksheet. Target worksheet.
'@param DisplayZeros Optional Boolean. When True, shows zero values. Defaults to False.
Private Sub RemoveGridLines(ByVal sh As Worksheet, Optional DisplayZeros As Boolean = False)
    Dim View As WorksheetView

    'The views are compared as objects, not by name: reading two sheet names per
    'view was two crossings into Excel where the identity test is one
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet Is sh Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next View
End Sub

'@label:prepare-print-sheet
'@sub-title Configure page setup for printing a worksheet
'@param sh Worksheet. Target worksheet to configure.
Private Sub PreparePrintSheet(ByVal sh As Worksheet)
    Application.PrintCommunication = False
    'Avoid printing rows and column number'
    With sh.PageSetup
        'Specifies the margins
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0.31)
        .FooterMargin = Application.InchesToPoints(0.31)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintTitleRows = "$5:$8" 'Those are rows to always keep on title
        .PrintTitleColumns = ""
        .PrintComments = xlPrintNoComments
        .PrintNotes = False
        'The quality of the print
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        'Landscape and paper size
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .FirstPageNumber = xlAutomatic
        .ORDER = xlDownThenOver
        .BlackAndWhite = False
        'Print the whole area and fit all columns in the worksheet
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        'Print Errors to blanks
        .PrintErrors = xlPrintErrorsBlank
    End With
    'Setting PrintCommunication to True can generate errors on MacOS
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

'@section Formatting
'===============================================================================
'Core formatting dispatcher.

'@label:apply-format
'@sub-title Apply a design-driven format to an Excel object
'@details
'Reads style metadata from the format table and applies it to the
'supplied object according to the given FormatScope constant.
'The object is sorted ONCE here, by its type, and then handed to the sub that
'formats that kind of object. Every branch used to open with its own
'TypeName test and its own Set, 31 times over. An object of the wrong kind
'for the requested scope still leaves quietly - callers rely on that silence.
'@param formattingObject Variant. The Excel object to format.
'@param scope Byte. FormatScope constant selecting the visual pattern.
'@param nbDecimals Integer. Decimal count for numeric formatting. Defaults to 0.
'@export
Public Sub ApplyFormat(ByVal formattingObject As Variant, _
                        ByVal scope As Byte, _
                        Optional ByVal nbDecimals As Integer = 0)

    Dim rng As Range
    Dim Char As Characters
    Dim sh As Worksheet
    Dim shp As Shape
    Dim tableObject As ListObject

    EnsureFormatReady

    Select Case TypeName(formattingObject)
    Case "Range"
        Set rng = formattingObject
    Case "Worksheet"
        Set sh = formattingObject
    Case "Characters"
        Set Char = formattingObject
    Case "Shape"
        Set shp = formattingObject
    Case "ListObject"
        Set tableObject = formattingObject
    Case Else
        Exit Sub
    End Select

    Select Case scope

    Case AllAnalysisSheet, AllLinelistSheet, HListCRFSheet, HListPrintedSheet, _
         HListCustomPivotTables
        If sh Is Nothing Then Exit Sub
        ApplyWorksheetFormat sh, scope

    Case HListMainLab, HListSublab, VListMainLab, VListSublab
        If Char Is Nothing Then Exit Sub
        ApplyCharactersFormat Char, scope

    Case LinelistSmallButton, AdminLargeButton, LinelistGeoButton
        If shp Is Nothing Then Exit Sub
        ApplyShapeFormat shp, scope

    Case HListTable, HListPrintedTable
        If tableObject Is Nothing Then Exit Sub
        ApplyListObjectFormat tableObject, scope

    Case Else
        'Every remaining scope formats a Range. A scope with no branch at all
        'falls through the inner Select Case and does nothing, as before
        If rng Is Nothing Then Exit Sub
        ApplyRangeFormat rng, scope, nbDecimals
    End Select
End Sub

'@label:apply-worksheet-format
'@sub-title Apply the worksheet-level format scopes
'@param sh Worksheet. The worksheet to format.
'@param scope Byte. FormatScope constant.
Private Sub ApplyWorksheetFormat(ByVal sh As Worksheet, ByVal scope As Byte)

    Dim styleValue As Variant
    Dim dimensionValue As Variant

    Select Case scope

    Case AllAnalysisSheet, AllLinelistSheet

        'Zeros are shown on both kinds of sheet
        RemoveGridLines sh, DisplayZeros:=True

        If scope = AllAnalysisSheet Then
            styleValue = AnalysisBaseFontSize()
            dimensionValue = DesignValue("default analysis column width", False)
            sh.Rows("2:3").RowHeight = 25
        Else
            styleValue = DesignValue("linelist base font size", False)
            dimensionValue = DesignValue("default linelist column width", False)
            sh.Rows("1:4").RowHeight = 25
        End If

        styleValue = IIf(styleValue = 0, 9, styleValue)
        dimensionValue = IIf(dimensionValue = 0, 25, dimensionValue)
        sh.Cells.Font.Size = styleValue
        sh.Cells.EntireColumn.ColumnWidth = dimensionValue

        'Column widths for dedicated sections
        If scope = AllAnalysisSheet Then
            sh.Columns(3).EntireColumn.ColumnWidth = dimensionValue + 5
            sh.Columns("A:B").EntireColumn.ColumnWidth = dimensionValue
        End If

    Case HListCRFSheet

        RemoveGridLines sh, DisplayZeros:=True
        sh.Cells.Font.Size = 9
        sh.Cells.EntireColumn.ColumnWidth = 4
        sh.Cells.EntireRow.RowHeight = 11
        sh.Range("A1").EntireColumn.ColumnWidth = 30

    Case HListPrintedSheet

        RemoveGridLines sh, DisplayZeros:=True
        sh.Cells.EntireRow.RowHeight = 20
        sh.Rows("1:4").EntireRow.RowHeight = 15
        PreparePrintSheet sh

    Case HListCustomPivotTables

        If Not IsEmpty(sh.Cells(1, 1).Value) And (sh.Cells(1, 2).Value <> "formated") Then
            RemoveGridLines sh, DisplayZeros:=True
            sh.Cells.EntireRow.RowHeight = 20
            sh.Visible = xlSheetVisible
            sh.Cells(1, 2).Value = "formated"
            sh.Cells(1, 2).Font.color = vbWhite
        ElseIf IsEmpty(sh.Cells(1, 1).Value) Then
            sh.Visible = xlSheetVeryHidden
        End If

    End Select
End Sub

'@label:apply-characters-format
'@sub-title Apply the label format scopes to a Characters object
'@param Char Characters. The label text to format.
'@param scope Byte. FormatScope constant.
Private Sub ApplyCharactersFormat(ByVal Char As Characters, ByVal scope As Byte)

    Dim fontColorValue As Variant
    Dim dimensionValue As Variant

    Select Case scope

    Case HListMainLab
        fontColorValue = DesignValue("hlist main label font color")
        dimensionValue = DesignValue("hlist main label font size", False)
        FormatCharacters Char:=Char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=True

    Case HListSublab
        fontColorValue = DesignValue("hlist sub label font color")
        dimensionValue = DesignValue("hlist sub label font size", False)
        FormatCharacters Char:=Char, FontSize:=dimensionValue, fontColor:=fontColorValue

    Case VListMainLab
        fontColorValue = DesignValue("vlist main label font color")
        dimensionValue = DesignValue("vlist main label font size", False)
        FormatCharacters Char:=Char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=False

    Case VListSublab
        fontColorValue = DesignValue("vlist sub label font color")
        dimensionValue = DesignValue("vlist sub label font size", False)
        FormatCharacters Char:=Char, FontSize:=dimensionValue, fontColor:=fontColorValue

    End Select
End Sub

'@label:apply-shape-format
'@sub-title Apply the button format scopes to a Shape
'@param shp Shape. The button to format.
'@param scope Byte. FormatScope constant.
Private Sub ApplyShapeFormat(ByVal shp As Shape, ByVal scope As Byte)

    Dim fontColorValue As Variant
    Dim interiorColorValue As Variant
    Dim dimensionValue As Variant

    Select Case scope

    Case LinelistSmallButton, AdminLargeButton

        fontColorValue = DesignValue("button default font color")
        interiorColorValue = DesignValue("button default interior color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue)
        dimensionValue = IIf(scope = AdminLargeButton, dimensionValue + 3, dimensionValue)

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, _
                    FontSize:=dimensionValue

    Case LinelistGeoButton

        fontColorValue = DesignValue("geo font color")
        interiorColorValue = DesignValue("geo interior color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue) + 1

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, _
                    FontSize:=dimensionValue

    End Select
End Sub

'@label:apply-listobject-format
'@sub-title Apply the entry table format scopes to a ListObject
'@param tableObject ListObject. The entry table to format.
'@param scope Byte. FormatScope constant.
Private Sub ApplyListObjectFormat(ByVal tableObject As ListObject, ByVal scope As Byte)

    Dim interiorColorValue As Variant
    Dim bordersColorValue As Variant

    interiorColorValue = DesignValue("entry table style", False)
    bordersColorValue = DesignValue("hlist table borders color")

    If interiorColorValue = "None, with borders" Then

        tableObject.TableStyle = ""
        DrawLines tableObject.DataBodyRange, At:="AvoidLeft", color:=bordersColorValue

    ElseIf interiorColorValue = "None" Then
        tableObject.TableStyle = ""
    Else
        'The design cell can name a style this Excel does not know; fall back to
        'a built-in one first so the table is never left unstyled
        On Error Resume Next
        tableObject.TableStyle = "TableStyleLight16"
        tableObject.TableStyle = interiorColorValue
        On Error GoTo 0
    End If

    'Increase the RowHeight of the ListObject on printed sheet
    If scope = HListPrintedTable Then
        tableObject.DataBodyRange.EntireRow.RowHeight = 70
    End If
End Sub

'@label:format-missing-look
'@sub-title The shared look of the missing and one-cell analysis scopes
'@details
'Six scopes share this block: the same two design colours, the same
'FormatRange call apart from the horizontal alignment, and the same double
'frame around the cells.
'@param rng Range. Target range.
'@param horizAlign Long. Horizontal alignment constant.
'@param drawInsideLines Boolean. True to rule the inside of the block as well.
Private Sub FormatMissingLook(ByVal rng As Range, _
                              ByVal horizAlign As Long, _
                              ByVal drawInsideLines As Boolean)

    Dim fontColorValue As Variant
    Dim interiorColorValue As Variant

    fontColorValue = DesignValue("missing font color")
    interiorColorValue = DesignValue("missing interior color")

    FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                FontSize:=AnalysisBaseFontSize(), isBold:=True, Horiz:=horizAlign

    'Only the missing row and column scopes rule the inside of the block
    If drawInsideLines Then DrawLines rng, color:=DesignValue("analysis table borders color")

    'A double frame on all four edges, at the hairline weight DrawLines defaults to
    DrawBorder rng, weight:=xlHairline, color:=fontColorValue, line:=xlDouble
End Sub

'@label:format-section-body
'@sub-title The shared tail of the section and sub-section scopes
'@param rng Range. Target range.
'@param interiorColorValue Variant. Fill colour.
'@param fontColorValue Variant. Font colour.
'@param fontSizeValue Variant. Font size in points.
Private Sub FormatSectionBody(ByVal rng As Range, _
                              ByVal interiorColorValue As Variant, _
                              ByVal fontColorValue As Variant, _
                              ByVal fontSizeValue As Variant)

    FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, _
                FontSize:=fontSizeValue, isBold:=True

    'Locking and merging both refuse on a protected sheet or on cells that are
    'already part of another merge. The colours are written either way, so a
    'failure here must not stop the format
    On Error Resume Next
    rng.Locked = True
    rng.Merge
    On Error GoTo 0
End Sub

'@label:apply-range-format
'@sub-title Apply every range-level format scope
'@details
'baseRng always holds the range the caller passed. rng is the working range
'and several scopes move it, so anything that has to measure from the
'original reads baseRng.
'@param baseRng Range. The range the caller passed.
'@param scope Byte. FormatScope constant.
'@param nbDecimals Integer. Decimal count for the CRF numeric scopes.
Private Sub ApplyRangeFormat(ByVal baseRng As Range, _
                             ByVal scope As Byte, _
                             ByVal nbDecimals As Integer)

    Dim rng As Range
    Dim cellRng As Range
    Dim sh As Worksheet

    Dim fontColorValue As Variant
    Dim interiorColorValue As Variant
    Dim bordersColorValue As Variant
    Dim styleValue As Variant
    Dim dimensionValue As Variant
    Dim verticalAlign As Long                    'Vertical Alignment
    Dim horizAlign As Long

    Set rng = baseRng

    Select Case scope

    Case AnalysisSection

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("table sections font color")
        dimensionValue = DesignValue("default analysis column width", False)

        Set sh = rng.Parent
        With sh
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 7))
        End With

        'Section color
        FormatRange rng, FontSize:=styleValue + 5, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, Verti:=xlVAlignBottom, _
                    isBold:=True
        'Draw a border arround the section
        DrawLines rng:=rng, weight:=xlMedium, color:=fontColorValue, At:="Bottom"

        'Gap after section
        On Error Resume Next
        rng.Offset(1).EntireRow.RowHeight = 45
        rng.Offset(1).EntireRow.RowHeight = dimensionValue
        On Error GoTo 0
        rng.Locked = False
        rng.Cells(1, 1).WrapText = True

    Case AnalysisTableTitle

        styleValue = AnalysisBaseFontSize()
        dimensionValue = DesignValue("gap for graphs", False)
        fontColorValue = DesignValue("table title font color")

        FormatRange rng, FontSize:=styleValue, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, _
                    Verti:=xlVAlignBottom
        rng.Font.Italic = True

        On Error Resume Next
        'Gap for graphs
        rng.EntireRow.RowHeight = 50
        rng.EntireRow.RowHeight = dimensionValue
        On Error GoTo 0

    Case HListCRFSection

        fontColorValue = DesignValue("sub section font color")
        dimensionValue = DesignValue("main section font size", False)
        dimensionValue = IIf(dimensionValue = 0, 12, dimensionValue)

        FormatRange rng, FontSize:=dimensionValue, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, Verti:=xlVAlignBottom, _
                    isBold:=True
        If (rng.Cells(1, 1).Value <> vbNullString) Then
            DrawLines rng, color:=fontColorValue, At:="Bottom", weight:=xlMedium
            rng.EntireRow.RowHeight = 14
        End If

    Case HListCRFSubSection

        fontColorValue = DesignValue("sub section font color")
        dimensionValue = DesignValue("sub section font size", False)
        dimensionValue = IIf(dimensionValue = 0, 11, dimensionValue)

        FormatRange rng, FontSize:=dimensionValue, _
                    fontColor:=fontColorValue, Horiz:=xlHAlignLeft, Verti:=xlVAlignBottom, _
                    isBold:=False

        If (rng.Cells(1, 1).Value <> vbNullString) Then
            DrawLines rng, color:=fontColorValue, At:="Bottom", weight:=xlThin
            rng.EntireRow.RowHeight = 14
        End If

    Case CRFText

        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 2), baseRng.Cells(1, 9))
        fontColorValue = DesignValue("main section interior color")
        DrawLines rng, color:=fontColorValue, At:="Bottom"

    Case CRFDate

        fontColorValue = DesignValue("main section interior color")
        'Format the overall range
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 2), baseRng.Cells(1, 9))
        DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue

        'Day
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 2), baseRng.Cells(1, 3))
        DrawLines rng, At:="AvoidTop", color:=fontColorValue
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue

        'Month
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 4), baseRng.Cells(1, 5))
        DrawLines rng, At:="AvoidTop"
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue

        'Year
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 6), baseRng.Cells(1, 9))
        DrawLines rng, At:="AvoidTop"
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue
        DrawLines rng, At:="Right", weight:=xlMedium, color:=fontColorValue

    Case CRFInteger

        fontColorValue = DesignValue("sub section font color")

        'Draw 3 squares
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 2), baseRng.Cells(1, 4))
        DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue
        DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue

    Case CRFDecimal, CRFPercentage

        fontColorValue = DesignValue("sub section font color")

        'Draw 3 squares
        Set rng = baseRng.Parent.Range(baseRng.Cells(1, 2), baseRng.Cells(1, 7))
        DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue
        DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue

        'Draw Decimals
        FormatRange rng.Cells(1, 3), fillValue:=".", Horiz:=xlHAlignCenter, _
                    Verti:=xlVAlignCenter, FontSize:=16, isBold:=True

        'Draw the cases for the number of decimals
        If nbDecimals > 0 Then
            Set rng = rng.Parent.Range(rng.Cells(1, 4), rng.Cells(1, 4 + nbDecimals))
            DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue
            DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue
        End If

        'Add percentage if required
        If scope = CRFPercentage Then
           FormatRange rng.Cells(1, 5 + nbDecimals), fillValue:="%", Horiz:=xlHAlignCenter, _
                    Verti:=xlVAlignCenter, FontSize:=16, isBold:=True
        End If

    Case CRFChoiceHeader

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("sub section font color")
        FormatRange rng, fontColor:=fontColorValue, isBold:=False, _
                         Verti:=xlVAlignCenter, FontSize:=styleValue - 2

    Case CRFChoiceValues

        fontColorValue = DesignValue("sub section font color")
        DrawBorder rng, color:=fontColorValue

    Case CRFFormatCustom

        fontColorValue = DesignValue("sub section font color")

        'Draw 5 squares
        Set rng = rng.Parent.Range(rng.Cells(1, 2), rng.Cells(1, 6))
        DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue

    Case LinelistSelectionDropdown

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("select dropdown font color")
        interiorColorValue = DesignValue("dropdown interior color")
        FormatRange rng, fontColor:=fontColorValue, _
                    interiorColor:=interiorColorValue, _
                    isBold:=False, FontSize:=styleValue - 2, _
                    Verti:=xlVAlignCenter, wrap:="wrap"
        DrawLines rng:=rng, weight:=xlThin, color:=vbWhite
        rng.Locked = False

    Case AnalysisCategoriesLabelRow, AnalysisCategoriesLabelColumn, AnalysisUnivariateHeader

        styleValue = AnalysisBaseFontSize()
        verticalAlign = IIf(scope = AnalysisCategoriesLabelRow, xlVAlignCenter, xlVAlignBottom)
        horizAlign = IIf(scope = AnalysisUnivariateHeader, xlHAlignCenter, xlHAlignLeft)

        fontColorValue = DesignValue("categories names font color")
        FormatRange rng, fontColor:=fontColorValue, isBold:=True, FontSize:=styleValue - 1, _
                    Horiz:=horizAlign, Verti:=verticalAlign, wrap:="wrap"

    Case AnalysisCategoriesUnivariateRow, AnalysisFirstRow, AnalysisCategoriesBivariateRow, _
         AnalysisCategoriesBivariateColumn

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("table categories font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("table categories interior color")

        horizAlign = IIf(scope = AnalysisCategoriesBivariateColumn, xlHAlignCenter, xlHAlignLeft)

        'Format the categories
        If scope <> AnalysisFirstRow Then
            FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue, _
                        Horiz:=horizAlign, Verti:=xlVAlignCenter, interiorColor:=interiorColorValue
            DrawLines rng, color:=bordersColorValue
        Else
            FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue - 2, _
                        Horiz:=xlHAlignCenter, Verti:=xlVAlignCenter
            DrawLines rng, color:=bordersColorValue
        End If

        If scope = AnalysisCategoriesBivariateColumn Or scope = AnalysisFirstRow Then
            DrawLines rng, At:="Left", color:=bordersColorValue, line:=xlDouble
            DrawLines rng, At:="Right", color:=bordersColorValue, line:=xlDouble
        Else
            DrawBorder rng, color:=bordersColorValue
        End If

        'Borders at top and bottom for bivariate
        If scope = AnalysisCategoriesBivariateRow Then
            DrawLines rng, line:=xlDouble, At:="Top", color:=bordersColorValue
            DrawLines rng, line:=xlDouble, At:="Bottom", color:=bordersColorValue
        End If

        'Wrap text on column and rows if too long
        rng.WrapText = True

    Case AnalysisGlobalSummaryRow

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("table categories font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("table categories interior color")

        Set sh = rng.Parent
        FormatRange rng, fontColor:=fontColorValue, FontSize:=styleValue, _
                    Horiz:=xlHAlignLeft, Verti:=xlVAlignCenter, interiorColor:=interiorColorValue
        'All the range line
        With sh
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 3))
            DrawLines rng, color:=bordersColorValue, weight:=xlThin
            DrawLines rng, color:=bordersColorValue, At:="Top"
            DrawLines rng, color:=bordersColorValue, At:="Bottom"
            Set rng = .Range(rng.Cells(1, 2), rng.Cells(1, 3))
            FormatRange rng, FontSize:=styleValue - 2, Horiz:=xlHAlignRight
        End With

    Case AnalysisInteriorTable, AnalysisWholeTable

        bordersColorValue = DesignValue("analysis table borders color")

        If scope = AnalysisInteriorTable Then
            DrawLines rng, color:=bordersColorValue
            'The four edges, all double, at the hairline weight DrawLines defaults to
            DrawBorder rng, weight:=xlHairline, color:=bordersColorValue, line:=xlDouble
        End If

        If scope = AnalysisWholeTable Then DrawBorder rng, color:=bordersColorValue

    Case AnalysisTimeSeriesHeader

        styleValue = AnalysisBaseFontSize()
        fontColorValue = DesignValue("time series header font")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("time series header interior")

        On Error Resume Next
        rng.Merge
        On Error GoTo 0
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue + 1, Horiz:=xlHAlignCenter, Verti:=xlVAlignCenter, _
                    isBold:=True, wrap:="wrap"
        DrawBorder rng, weight:=xlHairline, color:=bordersColorValue, line:=xlDouble

    Case AnalysisMissingRow, AnalysisMissingColumn, AnalysisMissingRowWithoutHeader

        FormatMissingLook rng, xlHAlignRight, drawInsideLines:=True

        If scope = AnalysisMissingRow Then
            'Straight to the range formatter: the object kind and the base font
            'size were both settled on the way in
            ApplyRangeFormat baseRng.Cells(1, 0), AnalysisOneCellLeft, 0
        End If

    Case AnalysisOneCell, AnalysisOneCellLeft, AnalysisOneCellCenter

        horizAlign = Switch(scope = AnalysisOneCellLeft, xlHAlignLeft, _
                            scope = AnalysisOneCellCenter, xlHAlignCenter, _
                            True, xlHAlignRight)

        FormatMissingLook rng, horizAlign, drawInsideLines:=False

    Case AnalysisPercent

        FormatRange rng, NumFormat:="0.00%"

    Case LinelistHiddenCell

        rng.Interior.color = vbWhite
        rng.Font.color = vbWhite
        rng.Locked = True
        rng.FormulaHidden = True

    Case AnalysisEndTable

        dimensionValue = DesignValue("gap for graphs", False)

        On Error Resume Next

        'Gap for graphs
        rng.EntireRow.RowHeight = 50
        rng.Offset(1).EntireRow.RowHeight = 50

        rng.EntireRow.RowHeight = dimensionValue
        rng.Offset(1).EntireRow.RowHeight = dimensionValue

        On Error GoTo 0

    Case AnalysisInfoCell, AnalysisInfoCellLeft, LinelistWarning

        'Horizontal alignment to left for cellLeft
        horizAlign = IIf(scope = AnalysisInfoCellLeft, xlHAlignLeft, xlHAlignCenter)

        fontColorValue = Switch(scope = LinelistWarning, vbRed, IsEmpty(rng), "", _
                            True, DesignValue("missing font color"))

        FormatRange rng, fontColor:=fontColorValue, _
                    FontSize:=Switch(scope = AnalysisInfoCellLeft, 11, scope = LinelistWarning, 12, True, 8), _
                    isBold:=(scope = AnalysisInfoCellLeft Or scope = LinelistWarning), _
                    Horiz:=horizAlign, Verti:=xlVAlignCenter, wrap:="wrap"
        If scope <> LinelistWarning Then rng.EntireRow.RowHeight = 45

    Case HListCalculatedFormulaCell

        fontColorValue = DesignValue("calculated formula font color")
        interiorColorValue = DesignValue("calculated formula column interior color")
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue
        rng.Locked = True
        rng.FormulaHidden = True

    Case HListCalculatedFormulaHeader

        interiorColorValue = DesignValue("calculated formula header color")
        FormatRange rng, interiorColor:=interiorColorValue, Verti:=xlVAlignCenter, _
                    Horiz:=xlHAlignLeft

    Case HListGeoHeader, HListGeo

        If scope = HListGeoHeader Then
            interiorColorValue = DesignValue("geo interior color")
        Else
            'vbNullString tells FormatRange to leave the fill alone
            interiorColorValue = vbNullString
        End If
        fontColorValue = DesignValue("geo font color")
        FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, _
                    Verti:=xlVAlignCenter, Horiz:=xlHAlignLeft

    Case HListSimpleHeaderCell

        FormatRange rng, Verti:=xlVAlignCenter, Horiz:=xlHAlignLeft, wrap:="wrap"

    Case HListSubSection, VListSubSection

        Set sh = rng.Parent
        'Borders arround subsections
        bordersColorValue = DesignValue("hlist and vlist table borders color")

        With sh
            If scope = HListSubSection Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(4, rng.Columns.Count))
            Else
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 3))
            End If
            DrawLines rng, color:=bordersColorValue

            'Draw borders for the range
            If scope = HListSubSection Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(4, rng.Columns.Count))
                DrawBorder rng, color:=bordersColorValue
            End If
        End With

        'format the range.
        interiorColorValue = DesignValue("sub section interior color")
        fontColorValue = DesignValue("sub section font color")
        dimensionValue = DesignValue("sub section font size", False)
        'Same fallback as every sibling scope; it was lost here because the
        'assignment above it wrote 11 and the design read overwrote it
        dimensionValue = IIf(dimensionValue = 0, 11, dimensionValue)

        FormatSectionBody baseRng, interiorColorValue, fontColorValue, dimensionValue

    Case HListSection, VListSection

        'draw borders
        Set sh = rng.Parent
        bordersColorValue = DesignValue("hlist and vlist table borders color")

        With sh

            If scope = HListSection Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(5, rng.Columns.Count))
                DrawBorder rng, color:=bordersColorValue, weight:=xlMedium
                DrawLines rng, color:=bordersColorValue, At:="Bottom"
            Else
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 4))
                DrawBorder rng, color:=bordersColorValue, weight:=xlThin
            End If
        End With

        'format the range
        fontColorValue = DesignValue("main section font color")
        interiorColorValue = DesignValue("main section interior color")
        dimensionValue = DesignValue("main section font size", False)
        dimensionValue = IIf(dimensionValue = 0, 12, dimensionValue)

        FormatSectionBody baseRng, interiorColorValue, fontColorValue, dimensionValue

    Case AnalysisDateCell

        rng.NumberFormat = "dd/mm/yyyy"

    Case HListHeader, HListPrintedHeader

        dimensionValue = IIf(scope = HListPrintedHeader, 5, 1)
        interiorColorValue = DesignValue("hlist table header color")
        FormatRange rng, fontColor:=interiorColorValue, _
                    interiorColor:=interiorColorValue, _
                    FontSize:=dimensionValue
        rng.Locked = False

        If scope = HListPrintedHeader Then

            'hide the headers
            rng.EntireRow.HIDDEN = True

            'Now formatting the headers (main label)
            dimensionValue = DesignValue("hlist main label font size", False)
            dimensionValue = IIf(dimensionValue = 0, 12, dimensionValue + 3)
            Set rng = rng.Offset(-1)
            FormatRange rng, FontSize:=dimensionValue
            rng.EntireRow.RowHeight = 100

            'AutoFit columns
            For Each cellRng In rng
                If Not cellRng.EntireColumn.HIDDEN Then
                    cellRng.EntireColumn.ColumnWidth = 30
                    cellRng.WrapText = True
                End If
            Next cellRng

            'Formating sub section (increasing the font)
            dimensionValue = DesignValue("sub section font size", False)
            dimensionValue = IIf(dimensionValue = 0, 18, dimensionValue + 7)
            Set rng = rng.Offset(-1)
            FormatRange rng, FontSize:=dimensionValue

            'Formating main section (increasing the font)
            dimensionValue = DesignValue("main section font size", False)
            dimensionValue = IIf(dimensionValue = 0, 22, dimensionValue + 10)
            Set rng = rng.Offset(-1)
            FormatRange rng, FontSize:=dimensionValue
        End If

    End Select

End Sub

'@section Import
'===============================================================================
'Import formatting definitions from another worksheet.

'@label:import
'@sub-title Import formatting definitions from another worksheet
'@details
'Reads design columns from the source worksheet and merges them into
'the current format table via CustomTable.Import. Updates the active
'design name from the source DESIGNTYPE named range.
'@param inpsh Worksheet. Source worksheet containing the format table.
'@export
Public Sub Import(ByVal inpsh As Worksheet)

    Dim targetSheet As Worksheet
    Dim inputTable As ListObject
    Dim targetTable As ListObject
    Dim inputTab As CustomTable
    Dim targetTab As CustomTable
    Dim designList As BetterArray
    Dim sourceDesign As Range

    If inpsh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Import worksheet is required"
    End If

    If inpsh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Import worksheet must expose a format table"
    End If

    Set sourceDesign = NamedRange(inpsh, "DESIGNTYPE")
    If sourceDesign Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Range 'DESIGNTYPE' is missing from import worksheet"
    End If

    EnsureFormatReady
    Set targetSheet = Wksh()
    Set targetTable = FormatTable()
    Set inputTable = inpsh.ListObjects(1)

    'Import the design names. The field is written straight, not through the
    'DesignName property: that property calls GuardNotSealed and Create seals
    'every instance, so the property route would raise on every import
    targetSheet.Range("DESIGNTYPE").Value = sourceDesign.Value
    this.design = CStr(targetSheet.Range("DESIGNTYPE").Value)

    'Import the different formats
    Set designList = New BetterArray
    designList.LowerBound = 1
    designList.Push DEFAULT_DESIGN, DESIGN_TWO, DESIGN_USER_DEFINED
    Set targetTab = CustomTable.Create(targetTable)
    Set inputTab = CustomTable.Create(inputTable)

    targetTab.Import inputTab, strictColumnSearch:=True, formatHeaders:=designList

    ResetCaches
End Sub


'@section Export
'===============================================================================
'Export the LL format worksheet into another workbook.

'@label:export
'@sub-title Export the format table to another workbook
'@details
'Copies the format table and cell styles to a new worksheet in the
'destination workbook. When a worksheet with the same name already
'exists, imports from it instead of exporting.
'@param toWkb Workbook. Destination workbook.
'@export
Public Sub Export(ByVal toWkb As Workbook)

    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As ListObject
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject
    Dim exportTable As CustomTable
    Dim sheetName As String

    If toWkb Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Export workbook is required"
    End If

    EnsureFormatReady
    Set sourceSheet = Wksh()
    sheetName = sourceSheet.Name
    Set targetWorkbook = toWkb

    Set targetSheet = ResolveWorksheet(targetWorkbook, sheetName)
    If Not targetSheet Is Nothing Then
        Import targetSheet
        Exit Sub
    End If

    Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
    targetSheet.Name = sheetName

    Set sourceTable = FormatTable()
    Set exportTable = CustomTable.Create(sourceTable)
    exportTable.Export targetSheet, addListObject:=True

    Set targetTable = ResolveFormatTable(targetSheet, sourceTable.Name)
    If targetTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Exported format table is missing on the destination worksheet"
    End If

    CopyFormatTableStyles sourceTable, targetTable
    EnsureDesignTypeRange sourceSheet, targetSheet, targetTable
End Sub

'@label:resolve-worksheet
'@sub-title Look up a worksheet by name without raising errors
'@return Worksheet. The found worksheet, or Nothing.
Private Function ResolveWorksheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    'Probe for the worksheet without raising errors when it is missing.
    On Error Resume Next
        Set candidate = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    Set ResolveWorksheet = candidate
End Function

'@label:resolve-format-table
'@sub-title Locate a ListObject on the target sheet by name or index
'@return ListObject. The found table, or Nothing.
Private Function ResolveFormatTable(ByVal targetSheet As Worksheet, ByVal preferredName As String) As ListObject
    Dim candidate As ListObject

    If targetSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set candidate = targetSheet.ListObjects(preferredName)
    On Error GoTo 0

    If candidate Is Nothing Then
        On Error Resume Next
            If targetSheet.ListObjects.Count > 0 Then
                Set candidate = targetSheet.ListObjects(1)
            End If
        On Error GoTo 0
    End If

    Set ResolveFormatTable = candidate
End Function

'@label:copy-format-table-styles
'@sub-title Copy cell-level styles between two format tables
'@details
'Column at a time. Excel answers Null for a property that is not the same on
'every cell of a range, so a column styled the same all the way down is
'copied with six writes instead of six writes per cell, and only a mixed block
'falls back to the cell loop. The header row is taken on its own, because it
'is styled differently from the body and would otherwise make every column
'mixed.
'What this CANNOT do: the design columns hold a different colour on every row,
'which is the whole point of the format table, so those columns are still
'walked cell by cell. VBA has no way to copy formats in a block off the
'clipboard. Copy plus PasteSpecial xlPasteFormats would do the whole table in
'one step, but the clipboard is the least reliable path on Mac Excel and this
'runs on the user's own workbook, so it is not used here.
Private Sub CopyFormatTableStyles(ByVal sourceTable As ListObject, ByVal targetTable As ListObject)
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim columnIndex As Long
    Dim rowCount As Long
    Dim columnCount As Long

    If sourceTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source format table is required for export"
    End If

    If targetTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Target format table is required for export"
    End If

    Set sourceRange = sourceTable.Range
    Set targetRange = targetTable.Range

    rowCount = sourceRange.Rows.Count
    columnCount = sourceRange.Columns.Count

    If rowCount <> targetRange.Rows.Count Or columnCount <> targetRange.Columns.Count Then
        ThrowError ProjectError.ErrorUnexpectedState, "Export format table dimensions do not match the source"
    End If

    'The header row goes across on its own. It is styled as one band on every
    'format sheet, and it is ALSO the reason a whole table column is never
    'uniform - mixing it in with the body would send every column to the cell
    'loop and the block path would never run at all
    CopyStyles sourceRange.Rows(1), targetRange.Rows(1)

    If rowCount <= 1 Then Exit Sub

    'Then one body column at a time. The text columns carry no per-row styling
    'and cross in one step each; the design columns really do hold a different
    'colour on every row, so those are the ones that fall to the cell loop
    For columnIndex = 1 To columnCount
        CopyStyles sourceRange.Columns(columnIndex).Offset(1).Resize(rowCount - 1), _
                   targetRange.Columns(columnIndex).Offset(1).Resize(rowCount - 1)
    Next columnIndex
End Sub

'@label:copy-styles
'@sub-title Copy the styles of one block, in one step when the block allows it
'@param sourceBlock Range. Block to read.
'@param targetBlock Range. Block of the same shape to write.
Private Sub CopyStyles(ByVal sourceBlock As Range, ByVal targetBlock As Range)
    If Not CopyStylesAsBlock(sourceBlock, targetBlock) Then
        CopyStylesByCell sourceBlock, targetBlock
    End If
End Sub

'@label:copy-styles-as-block
'@sub-title Copy a block of styles in a single write when the block is uniform
'@details
'Excel answers Null for a property that is not the same on every cell of a
'range, so reading the six properties off the whole block is both the test
'and the read.
'@param sourceBlock Range. Block to read.
'@param targetBlock Range. Block of the same shape to write.
'@return Boolean. True when the whole block was copied in one step.
Private Function CopyStylesAsBlock(ByVal sourceBlock As Range, _
                                   ByVal targetBlock As Range) As Boolean
    Dim patternValue As Variant
    Dim interiorColorValue As Variant
    Dim fontColorValue As Variant
    Dim boldValue As Variant
    Dim italicValue As Variant
    Dim sizeValue As Variant

    With sourceBlock
        patternValue = .Interior.Pattern
        interiorColorValue = .Interior.color
        fontColorValue = .Font.color
        boldValue = .Font.Bold
        italicValue = .Font.Italic
        sizeValue = .Font.Size
    End With

    'Null on any one of them means the block is mixed and cannot be copied whole
    If IsNull(patternValue) Or IsNull(interiorColorValue) Then Exit Function
    If IsNull(fontColorValue) Or IsNull(boldValue) Then Exit Function
    If IsNull(italicValue) Or IsNull(sizeValue) Then Exit Function

    With targetBlock
        .Interior.Pattern = patternValue
        .Interior.color = interiorColorValue
        .Font.color = fontColorValue
        .Font.Bold = boldValue
        .Font.Italic = italicValue
        .Font.Size = sizeValue
    End With

    CopyStylesAsBlock = True
End Function

'@label:copy-styles-by-cell
'@sub-title Copy a block of styles cell by cell
'@details
'Both blocks are walked with a single Cells index, so the same loop serves a
'one-row block and a one-column block. Each cell reads its Interior and its
'Font once instead of chaining through them six times.
'@param sourceBlock Range. Block to read.
'@param targetBlock Range. Block of the same shape to write.
Private Sub CopyStylesByCell(ByVal sourceBlock As Range, ByVal targetBlock As Range)
    Dim cellIndex As Long
    Dim cellCount As Long
    Dim patternValue As Variant
    Dim interiorColorValue As Variant
    Dim fontColorValue As Variant
    Dim boldValue As Variant
    Dim italicValue As Variant
    Dim sizeValue As Variant

    cellCount = sourceBlock.Cells.Count

    For cellIndex = 1 To cellCount
        With sourceBlock.Cells(cellIndex)
            patternValue = .Interior.Pattern
            interiorColorValue = .Interior.color
            fontColorValue = .Font.color
            boldValue = .Font.Bold
            italicValue = .Font.Italic
            sizeValue = .Font.Size
        End With

        With targetBlock.Cells(cellIndex)
            .Interior.Pattern = patternValue
            .Interior.color = interiorColorValue
            .Font.color = fontColorValue
            .Font.Bold = boldValue
            .Font.Italic = italicValue
            .Font.Size = sizeValue
        End With
    Next cellIndex
End Sub

'@label:ensure-design-type-range
'@sub-title Create the DESIGNTYPE named range on the target sheet
Private Sub EnsureDesignTypeRange(ByVal sourceSheet As Worksheet, _
                                  ByVal targetSheet As Worksheet, _
                                  ByVal targetTable As ListObject)

    Dim sourceDesignRange As Range
    Dim targetDesignRange As Range
    Dim designLabel As String
    Dim headerRow As Long
    Dim headerColumn As Long
    Dim targetColumn As Long

    If sourceSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source worksheet is required for design export"
    End If

    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target worksheet is required for design export"
    End If

    If targetTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target format table is required for design export"
    End If

    Set sourceDesignRange = NamedRange(sourceSheet, "DESIGNTYPE")

    If Not sourceDesignRange Is Nothing Then
        designLabel = CStr(sourceDesignRange.Value)
        If LenB(designLabel) = 0 Then designLabel = DEFAULT_DESIGN
        Set targetDesignRange = targetSheet.Cells(sourceDesignRange.Row, sourceDesignRange.Column)
    Else
        designLabel = DesignName()
        If LenB(designLabel) = 0 Then designLabel = DEFAULT_DESIGN

        headerRow = targetTable.HeaderRowRange.Row
        headerColumn = targetTable.Range.Column
        targetColumn = headerColumn + targetTable.ListColumns.Count + 2

        Set targetDesignRange = targetSheet.Cells(headerRow, targetColumn)
    End If

    targetDesignRange.Value = designLabel

    On Error Resume Next
        targetSheet.Names("DESIGNTYPE").Delete
    On Error GoTo 0
    targetSheet.Names.Add Name:="DESIGNTYPE", RefersTo:=targetDesignRange, Visible:=False
End Sub

'@label:named-range
'@sub-title Retrieve a worksheet-scoped named range, or Nothing when it is absent
'@param sh Worksheet. The worksheet to read.
'@param rngName String. The name to look up.
'@return Range. The named range, or Nothing.
Private Function NamedRange(ByVal sh As Worksheet, ByVal rngName As String) As Range
    On Error Resume Next
        Set NamedRange = sh.Range(rngName)
    On Error GoTo 0
End Function

'@label:throw-error
'@sub-title Raise a typed project error
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, errorMessage
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
