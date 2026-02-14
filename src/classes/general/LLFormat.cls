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
Implements ILLFormat
Private Const CLASSNAME As String = "LLFormat"

'@class LLFormat
'@description
'Reads design information from the LL format table and exposes helpers
'to apply those styles while caching lookups for performance. Supports
'import/export of format definitions across workbooks. Each FormatScope
'constant maps to a specific visual pattern applied to Worksheets,
'Ranges, Shapes, ListObjects, or Characters.
'@depends CustomTable, Checking, BetterArray, ProjectError
'@version 1.0 (2026-02-09)
'@author Yves Amevoin

Private Const DEFAULT_DESIGN As String = "design 1"
Private Const COLUMN_LABEL As String = "Label"
Private Const DESIGN_TWO As String = "design 2"
Private Const DESIGN_USER_DEFINED As String = "user defined"
Private Const FALLBACK_COLOR As Long = vbBlack
Private Const FALLBACK_VALUE As Long = 0

Private Type TLLFormat
    sheet As Worksheet
    design As String
    formatTable As ListObject
    labelColumnIndex As Long
    designColumnCache As BetterArray
    labelRowCache As BetterArray
    hasChecks As Boolean
    checks As IChecking
    checkCounter As Long
End Type

Private Type TFormatting
    fontColour As Long
    fillColour As Long
    fontSize As Long
    borderColour As Long
End Type

Private this As TLLFormat

'@section Cache Management
'===============================================================================
'Lazy-init and invalidation of internal lookup caches.

'@label:reset-caches
'@sub-title Invalidate all cached references
Private Sub ResetCaches()
    Set this.formatTable = Nothing
    Set this.designColumnCache = Nothing
    Set this.labelRowCache = Nothing
End Sub

'@label:ensure-format-ready
'@sub-title Ensure caches are initialised and the worksheet is set
Private Sub EnsureFormatReady()
    If this.sheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Format worksheet reference is missing"
    End If
    If this.designColumnCache Is Nothing Then
        Set this.designColumnCache = New BetterArray
        this.designColumnCache.LowerBound = 1
    End If
    If this.labelRowCache Is Nothing Then
        Set this.labelRowCache = New BetterArray
        this.labelRowCache.LowerBound = 1
    End If
End Sub

'@label:ensure-format-sheet-ready
'@sub-title Validate a worksheet as a valid format source
'@param sh Worksheet. The worksheet to validate.
'@param resolvedDesign String (ByRef). Falls back to DEFAULT_DESIGN when the requested design column is absent.
Private Sub EnsureFormatSheetReady(ByVal sh As Worksheet, ByRef resolvedDesign As String)
    
    Dim Lo As ListObject
    Dim header As Range

    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Format worksheet must be supplied"
    End If

    If sh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "No format table found on worksheet '" & sh.Name & "'"
    End If

    Set Lo = sh.ListObjects(1)
    On Error Resume Next
        Set header = Lo.HeaderRowRange
    On Error GoTo 0

    If header Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Format table on worksheet '" & sh.Name & "' is missing headers"
    End If

    If Lo.DataBodyRange Is Nothing Or Lo.DataBodyRange.Rows.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Format table '" & lo.Name & "' does not contain any data rows"
    End If

    If Not ColumnExists(header, COLUMN_LABEL) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Column '" & COLUMN_LABEL & "' is required in format table '" & lo.Name & "'"
    End If

    If Not ColumnExists(header, DEFAULT_DESIGN) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Default design '" & DEFAULT_DESIGN & "' was not found in format table '" & lo.Name & "'"
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
        If this.sheet.ListObjects.Count = 0 Then
            ThrowError ProjectError.ElementNotFound, "Format worksheet must contain at least one table"
        End If
        Set this.formatTable = this.sheet.ListObjects(1)
    End If
    Set FormatTable = this.formatTable
End Function

'@label:header-range
'@sub-title Retrieve the header row range of the format table
'@return Range. The header row.
Private Function HeaderRange() As Range
    Dim lo As ListObject
    Set lo = FormatTable()
    On Error Resume Next
        Set HeaderRange = lo.HeaderRowRange
    On Error GoTo 0
    If HeaderRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Format table header row is missing"
    End If
End Function

'@label:ensure-label-column
'@sub-title Cache the column index of the Label column
Private Sub EnsureLabelColumn()
    Dim header As Range
    If this.labelColumnIndex > 0 Then Exit Sub
    Set header = HeaderRange()
    this.labelColumnIndex = ResolveColumnIndex(header, COLUMN_LABEL)
End Sub

'@label:resolve-column-index
'@sub-title Find the column index of a named header
'@param header Range. The header row to search.
'@param columnName String. Column caption to locate.
'@return Long. The 1-based column index.
Private Function ResolveColumnIndex(ByVal header As Range, ByVal columnName As String) As Long
    Dim matchCell As Range
    On Error Resume Next
        Set matchCell = header.Find(What:=columnName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If matchCell Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Column '" & columnName & "' missing from format table"
    End If
    ResolveColumnIndex = matchCell.Column
End Function

'@label:design-column-index
'@sub-title Resolve and cache the column index for a design name
'@param designName String. Name of the design column.
'@return Long. The 1-based column index.
Private Function DesignColumnIndex(ByVal designName As String) As Long
    Dim idx As Long
    Dim entry As Variant
    Dim header As Range
    EnsureFormatReady

    For idx = this.designColumnCache.LowerBound To this.designColumnCache.UpperBound
        entry = this.designColumnCache.Item(idx)
        If IsArray(entry) Then
            'With lowerbound = 1 for the BetterArray, the inner stored array starts at 1
            If StrComp(CStr(entry(1)), designName, vbTextCompare) = 0 Then
                DesignColumnIndex = CLng(entry(2))
                Exit Function
            End If
        End If
    Next idx

    Set header = HeaderRange()
    DesignColumnIndex = ResolveColumnIndex(header, designName)
    this.designColumnCache.Push Array(designName, DesignColumnIndex)
End Function

'@label:label-row-index
'@sub-title Resolve and cache the row index for a format label
'@param labelName String. Label text to locate in the Label column.
'@return Long. The 1-based row index, or 0 when not found.
Private Function LabelRowIndex(ByVal labelName As String) As Long
    Dim idx As Long
    Dim entry As Variant
    Dim dataColumn As Range
    Dim matchCell As Range

    EnsureFormatReady
    EnsureLabelColumn

    For idx = this.labelRowCache.LowerBound To this.labelRowCache.UpperBound
        entry = this.labelRowCache.Item(idx)
        If IsArray(entry) Then
            'With lowerbound = 1 for the BetterArray, the stored array starts at 1
            If StrComp(CStr(entry(1)), labelName, vbTextCompare) = 0 Then
                LabelRowIndex = CLng(entry(2))
                Exit Function
            End If
        End If
    Next idx

    Set dataColumn = FormatTable().ListColumns(COLUMN_LABEL).DataBodyRange
    On Error Resume Next
        Set matchCell = dataColumn.Find(What:=labelName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If Not matchCell Is Nothing Then
        LabelRowIndex = matchCell.Row
        this.labelRowCache.Push Array(labelName, LabelRowIndex)
    End If
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
'@return ILLFormat. A fully initialised LLFormat instance.
'@export
Public Function Create(ByVal formatSh As Worksheet, _
                       Optional ByVal designType As String = DEFAULT_DESIGN) As ILLFormat

    Dim resolvedDesign As String

    resolvedDesign = designType

    EnsureFormatSheetReady formatSh, resolvedDesign

    With New LLFormat
        Set .Wksh = formatSh
        .DesignName = resolvedDesign
        Set Create = .Self
    End With

End Function

'@label:column-exists
'@sub-title Test whether a named column exists in a header range
'@param header Range. The header row to search.
'@param columnName String. Column caption to test.
'@return Boolean. True when the column is found.
Private Function ColumnExists(ByVal header As Range, ByVal columnName As String) As Boolean
    Dim found As Range
    On Error Resume Next
        Set found = header.Find(What:=columnName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    ColumnExists = Not (found Is Nothing)
End Function

'@section Public Interface
'===============================================================================
'Public properties exposed by the factory and the interface.

'@label:self
'@prop-title Retrieve this formatter as an interface
'@return ILLFormat. Self-reference.
Public Property Get Self() As ILLFormat
    Set Self = Me
End Property

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
'@return IChecking. The checking instance, or Nothing.
Public Property Get CheckingValues() As IChecking
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
'@param labelName String. Label text stored in the format table.
'@param returnColor Boolean. When True returns Interior.Color; otherwise the raw value. Defaults to True.
'@return Variant. The colour (Long) or cell value.
'@export
Private Function DesignValue(ByVal labelName As String, _
                             Optional ByVal returnColor As Boolean = True) As Variant

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetCell As Range
    Dim resultValue As Variant

    EnsureFormatReady

    rowIndex = LabelRowIndex(labelName)
    If rowIndex = 0 Then
        LogInfo "Label '" & labelName & "' not found; returning fallback value"
        DesignValue = IIf(returnColor, FALLBACK_COLOR, FALLBACK_VALUE)
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
End Function

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
    Dim borderpos As Byte
    Dim bordersList As BetterArray
    Dim counter As Long

    Set bordersList = New BetterArray

    Select Case At

    Case "All"
        bordersList.Push xlEdgeLeft, xlEdgeRight, xlEdgeBottom, xlEdgeTop, _ 
                         xlInsideHorizontal, xlInsideVertical
    Case "AvoidLeft", "AvoidRight"
        bordersList.Push xlEdgeBottom, xlEdgeTop, xlInsideHorizontal
    Case "AvoidTop", "AvoidBottom"
        bordersList.Push xlEdgeLeft, xlEdgeRight, xlInsideVertical
    Case "Left"
        bordersList.Push xlEdgeLeft
    Case "Right"
        bordersList.Push xlEdgeRight
    Case "Bottom"
        bordersList.Push xlEdgeBottom
    Case "Top"
        bordersList.Push xlEdgeTop
    Case Else
        bordersList.Push xlEdgeBottom
    End Select

    With rng
        For counter = bordersList.LowerBound To bordersList.UpperBound
            borderpos = bordersList.Item(counter)
            With .Borders(borderpos)
                .weight = weight
                .LineStyle = line
                .color = color
                .TintAndShade = 0.2
            End With
        Next
    End With
End Sub

'@label:draw-border
'@sub-title Draw a rectangular border around a range
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
            .LineStyle = line
            .color = color
            .TintAndShade = 0.2
            .weight = weight
        End With
    Next
End Sub

'@label:format-range
'@sub-title Apply fill, font, alignment, and number format to a range
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
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = sh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next
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
'@param formattingObject Variant. The Excel object to format.
'@param scope Byte. FormatScope constant selecting the visual pattern.
'@param nbDecimals Integer. Decimal count for numeric formatting. Defaults to 0.
'@export
Private Sub ApplyFormat(ByVal formattingObject As Variant, _
                        ByVal scope As Byte, _
                        Optional ByVal nbDecimals As Integer = 0)

    Dim rng As Range
    Dim cellRng As Range
    Dim Char As Characters
    Dim sh As Worksheet
    Dim shp As Shape
    Dim Lo As listObject

    Dim fontColorValue As Variant
    Dim interiorColorValue As Variant
    Dim bordersColorValue As Variant
    Dim styleValue As Variant
    Dim dimensionValue As Variant
    Dim verticalAlign As Long                    'Vertical Alignment
    Dim horizAlign As Long


    EnsureFormatReady

    'Formatting on analysis  part:
    styleValue = DesignValue("analysis base font size", False)
    styleValue = IIf(styleValue = 0, 11, styleValue)

    Select Case scope
    Case AllAnalysisSheet, AllLinelistSheet
        If (Not TypeName(formattingObject) = "Worksheet") Then Exit Sub

        Set sh = formattingObject

        If scope = AllAnalysisSheet Then
            RemoveGridLines sh, DisplayZeros:=True
            dimensionValue = DesignValue("default analysis column width", False)
            sh.Rows("2:3").RowHeight = 25
        Else
            'on Linelist worksheets, show zeros
            RemoveGridLines sh, DisplayZeros:=(scope = AllLinelistSheet)
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
    
        If (Not TypeName(formattingObject) = "Worksheet") Then Exit Sub
        Set sh = formattingObject
        RemoveGridLines sh, DisplayZeros:=True
        sh.Cells.Font.Size = 9
        sh.Cells.EntireColumn.ColumnWidth = 4
        sh.Cells.EntireRow.RowHeight = 11
        sh.Range("A1").EntireColumn.ColumnWidth = 30

    Case HListPrintedSheet
        If (Not TypeName(formattingObject) = "Worksheet") Then Exit Sub
        Set sh = formattingObject

        RemoveGridLines sh, DisplayZeros:=True
        sh.Cells.EntireRow.RowHeight = 20
        sh.Rows("1:4").EntireRow.RowHeight = 15
        PreparePrintSheet sh
    
    Case HListCustomPivotTables
        If (Not TypeName(formattingObject) = "Worksheet") Then Exit Sub
        
        Set sh = formattingObject

        If Not IsEmpty(sh.Cells(1, 1).Value) And (sh.Cells(1, 2).Value <> "formated") Then
            RemoveGridLines sh, DisplayZeros:=True
            sh.Cells.EntireRow.RowHeight = 20
            sh.Visible = xlSheetVisible
            sh.Cells(1, 2).Value = "formated"
            sh.Cells(1, 2).Font.color = vbWhite
        ElseIf IsEmpty(sh.Cells(1, 1).Value) Then
            sh.Visible = xlSheetVeryHidden
        End If

    Case AnalysisSection

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
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
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
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

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject

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

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        
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

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        Set rng = rng.Parent.Range(rng.Cells(1, 2), rng.Cells(1, 9))
        fontColorValue = DesignValue("main section interior color")
        DrawLines rng, color:=fontColorValue, At:="Bottom"

    Case CRFDate

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        fontColorValue = DesignValue("main section interior color")
        'Format the overall range
        Set rng = formattingObject.Parent.Range(formattingObject.Cells(1, 2), formattingObject.Cells(1, 9))
        DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue

        'Day

        Set rng = rng.Parent.Range(formattingObject.Cells(1, 2), formattingObject.Cells(1, 3))
        DrawLines rng, At:="AvoidTop", color:=fontColorValue
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue

        'Month
        Set rng = rng.Parent.Range(formattingObject.Cells(1, 4), formattingObject.Cells(1, 5))
        DrawLines rng, At:="AvoidTop"
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue

        'Year
        Set rng = rng.Parent.Range(formattingObject.Cells(1, 6), formattingObject.Cells(1, 9))
        DrawLines rng, At:="AvoidTop"
        DrawLines rng, At:="Left", weight:=xlMedium, color:=fontColorValue
        DrawLines rng, At:="Right", weight:=xlMedium, color:=fontColorValue

    Case CRFInteger

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        fontColorValue = DesignValue("sub section font color")

        'Draw 3 squares
        Set rng = formattingObject.Parent.Range(formattingObject.Cells(1, 2), formattingObject.Cells(1, 4))
        DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue
        DrawLines rng, At:="Bottom", weight:=xlMedium, color:=fontColorValue

    Case CRFDecimal, CRFPercentage

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        fontColorValue = DesignValue("sub section font color")
        
        'Draw 3 squares
        Set rng = formattingObject.Parent.Range(formattingObject.Cells(1, 2), formattingObject.Cells(1, 7))
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

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("sub section font color")
        FormatRange rng, fontColor:=fontColorValue, isBold:=False, _
                         Verti:=xlVAlignCenter, FontSize:=styleValue - 2  

    Case CRFChoiceValues

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("sub section font color")
        DrawBorder rng, color:=fontColorValue

    Case CRFFormatCustom

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("sub section font color")
        
        'Draw 5 squares
        Set rng = rng.Parent.Range(rng.Cells(1, 2), rng.Cells(1, 6))
        DrawLines rng, At:="AvoidTop", weight:=xlMedium, color:=fontColorValue


    Case LinelistSelectionDropdown
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("select dropdown font color")
        interiorColorValue = DesignValue("dropdown interior color")
        FormatRange rng, fontColor:=fontColorValue, _
                    interiorColor:=interiorColorValue, _
                    isBold:=False, FontSize:=styleValue - 2, _
                    Verti:=xlVAlignCenter, wrap:="wrap"
        DrawLines rng:=rng, weight:=xlThin, color:=vbWhite
        rng.Locked = False

    Case AnalysisCategoriesLabelRow, AnalysisCategoriesLabelColumn, AnalysisUnivariateHeader
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject

        verticalAlign = IIf(scope = AnalysisCategoriesLabelRow, xlVAlignCenter, xlVAlignBottom)
        horizAlign = IIf(scope = AnalysisUnivariateHeader, xlHAlignCenter, xlHAlignLeft)

        fontColorValue = DesignValue("categories names font color")
        FormatRange rng, fontColor:=fontColorValue, isBold:=True, FontSize:=styleValue - 1, _
                    Horiz:=horizAlign, Verti:=verticalAlign, wrap:="wrap"

    Case AnalysisCategoriesUnivariateRow, AnalysisFirstRow, AnalysisCategoriesBivariateRow, AnalysisCategoriesBivariateColumn

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
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
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
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

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        bordersColorValue = DesignValue("analysis table borders color")

        If scope = AnalysisInteriorTable Then
            DrawLines rng, color:=bordersColorValue
            DrawLines rng, color:=bordersColorValue, At:="Top", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Bottom", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Right", line:=xlDouble
            DrawLines rng, color:=bordersColorValue, At:="Left", line:=xlDouble
        End If

        If scope = AnalysisWholeTable Then DrawBorder rng, color:=bordersColorValue

    Case AnalysisTimeSeriesHeader

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("time series header font")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("time series header interior")

        On Error Resume Next
        rng.Merge
        On Error GoTo 0
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue + 1, Horiz:=xlHAlignCenter, Verti:=xlVAlignCenter, isBold:=True, _ 
                    wrap:="wrap"
        DrawLines rng, At:="Top", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Left", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Bottom", color:=bordersColorValue, line:=xlDouble
        DrawLines rng, At:="Right", color:=bordersColorValue, line:=xlDouble

    Case AnalysisMissingRow, AnalysisMissingColumn, AnalysisMissingRowWithoutHeader

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("missing font color")
        bordersColorValue = DesignValue("analysis table borders color")
        interiorColorValue = DesignValue("missing interior color")

        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue, isBold:=True, Horiz:=xlHAlignRight

        'Draw lines
        DrawLines rng, color:=bordersColorValue
        DrawLines rng, color:=fontColorValue, At:="Top", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Left", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Bottom", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Right", line:=xlDouble

        If scope = AnalysisMissingRow Then
            Set rng = rng.Cells(1, 0)
            ApplyFormat rng, scope:=AnalysisOneCellLeft
        End If

    Case AnalysisOneCell, AnalysisOneCellLeft, AnalysisOneCellCenter

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("missing font color")
        interiorColorValue = DesignValue("missing interior color")
        horizAlign = Switch(scope = AnalysisOneCellLeft, xlHAlignLeft, _
                            scope = AnalysisOneCellCenter, xlHAlignCenter, _
                            True, xlHAlignRight)

        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue, _
                    FontSize:=styleValue, isBold:=True, Horiz:=horizAlign

        DrawLines rng, color:=fontColorValue, At:="Bottom", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Top", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Left", line:=xlDouble
        DrawLines rng, color:=fontColorValue, At:="Right", line:=xlDouble

    Case AnalysisPercent

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        FormatRange rng, NumFormat:="0.00%"

    Case LinelistHiddenCell

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        Set rng = formattingObject
        rng.Interior.color = vbWhite
        rng.Font.color = vbWhite
        rng.Locked = True
        rng.FormulaHidden = True

    Case AnalysisEndTable

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        Set rng = formattingObject
        dimensionValue = DesignValue("gap for graphs", False)

        On Error Resume Next

        'Gap for graphs
        rng.EntireRow.RowHeight = 50
        rng.Offset(1).EntireRow.RowHeight = 50

        rng.EntireRow.RowHeight = dimensionValue
        rng.Offset(1).EntireRow.RowHeight = dimensionValue

        On Error GoTo 0

    Case AnalysisInfoCell, AnalysisInfoCellLeft, LinelistWarning

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        'Horizontal alignment to left for cellLeft
        horizAlign = IIf(scope = AnalysisInfoCellLeft, xlHAlignLeft, xlHAlignCenter)

        Set rng = formattingObject
        fontColorValue = Switch(scope = LinelistWarning, vbRed, IsEmpty(rng), "", _
                            True, DesignValue("missing font color"))

        FormatRange rng, fontColor:=fontColorValue, _
                    FontSize:=Switch(scope = AnalysisInfoCellLeft, 11, scope = LinelistWarning, 12, True, 8), _
                    isBold:=(scope = AnalysisInfoCellLeft Or scope = LinelistWarning), _
                    Horiz:=horizAlign, Verti:=xlVAlignCenter, wrap:="wrap"
        If scope <> LinelistWarning Then rng.EntireRow.RowHeight = 45

    Case HListMainLab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set Char = formattingObject
        fontColorValue = DesignValue("hlist main label font color")
        dimensionValue = DesignValue("hlist main label font size", False)
        FormatCharacters Char:=Char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=True

    Case HListSublab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set Char = formattingObject
        fontColorValue = DesignValue("hlist sub label font color")
        dimensionValue = DesignValue("hlist sub label font size", False)

        FormatCharacters Char:=Char, FontSize:=dimensionValue, fontColor:=fontColorValue

    Case HListCalculatedFormulaCell
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        fontColorValue = DesignValue("calculated formula font color")
        interiorColorValue = DesignValue("calculated formula column interior color")
        FormatRange rng, fontColor:=fontColorValue, interiorColor:=interiorColorValue
        rng.Locked = True
        rng.FormulaHidden = True

    Case HListCalculatedFormulaHeader
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        interiorColorValue = DesignValue("calculated formula header color")
        FormatRange rng, interiorColor:=interiorColorValue, Verti:=xlVAlignCenter, Horiz:=xlHAlignLeft

    Case HListGeoHeader, HListGeo
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        If scope = HListGeoHeader Then
            interiorColorValue = DesignValue("geo interior color")
        Else
            interiorColorValue = vbNullString
        End If
        fontColorValue = DesignValue("geo font color")
        FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, Verti:=xlVAlignCenter, Horiz:=xlHAlignLeft

    Case VListMainLab
        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set Char = formattingObject
        fontColorValue = DesignValue("vlist main label font color")
        dimensionValue = DesignValue("vlist main label font size", False)
        FormatCharacters Char:=Char, fontColor:=fontColorValue, _
                         FontSize:=dimensionValue, isBold:=False

    Case VListSublab

        If (Not TypeName(formattingObject) = "Characters") Then Exit Sub
        Set Char = formattingObject
        fontColorValue = DesignValue("vlist sub label font color")
        dimensionValue = DesignValue("vlist sub label font size", False)
        FormatCharacters Char:=Char, FontSize:=dimensionValue, fontColor:=fontColorValue

    Case HListSimpleHeaderCell

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
        FormatRange rng, Verti:=xlVAlignCenter, Horiz:=xlHAlignLeft, wrap:="wrap"

    Case LinelistSmallButton, AdminLargeButton

        If (Not TypeName(formattingObject) = "Shape") Then Exit Sub
        Set shp = formattingObject

        fontColorValue = DesignValue("button default font color")
        interiorColorValue = DesignValue("button default interior color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue)
        dimensionValue = IIf(scope = AdminLargeButton, dimensionValue + 3, dimensionValue)

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue

    Case LinelistGeoButton

        If (Not TypeName(formattingObject) = "Shape") Then Exit Sub
        Set shp = formattingObject

        fontColorValue = DesignValue("geo font color")
        interiorColorValue = DesignValue("geo interior color")
        dimensionValue = DesignValue("linelist base font size", False)
        dimensionValue = IIf(dimensionValue = 0, 9, dimensionValue) + 1

        FormatShape shp, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue

    Case HListSubSection, VListSubSection

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        Set rng = formattingObject
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
        Set rng = formattingObject

        interiorColorValue = DesignValue("sub section interior color")
        fontColorValue = DesignValue("sub section font color")
        dimensionValue = 11
        dimensionValue = DesignValue("sub section font size", False)

        FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue, _
                    isBold:=True

        On Error Resume Next
        rng.Locked = True
        rng.Merge
        On Error GoTo 0

    Case HListSection, VListSection

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        'draw borders
        Set rng = formattingObject
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
        Set rng = formattingObject

        fontColorValue = DesignValue("main section font color")
        interiorColorValue = DesignValue("main section interior color")
        dimensionValue = DesignValue("main section font size", False)
        dimensionValue = IIf(dimensionValue = 0, 12, dimensionValue)

        FormatRange rng, interiorColor:=interiorColorValue, fontColor:=fontColorValue, FontSize:=dimensionValue, _
                    isBold:=True

        On Error Resume Next
        rng.Locked = True
        rng.Merge
        On Error GoTo 0

    Case AnalysisDateCell
        If (Not TypeName(formattingObject) = "Range") Then Exit Sub
        'draw borders
        Set rng = formattingObject
        rng.NumberFormat = "dd/mm/yyyy"


    Case HListTable, HListPrintedTable

        If (Not TypeName(formattingObject) = "ListObject") Then Exit Sub
        Set Lo = formattingObject

        interiorColorValue = DesignValue("entry table style", False)
        bordersColorValue = DesignValue("hlist table borders color")

        If interiorColorValue = "None, with borders" Then
            
            Lo.TableStyle = ""
            DrawLines Lo.DataBodyRange, At:="AvoidLeft", color:=bordersColorValue

        ElseIf interiorColorValue = "None" Then
            Lo.TableStyle = ""
        Else
            On Error Resume Next
            Lo.TableStyle = "TableStyleLight16"
            Lo.TableStyle = interiorColorValue
            On Error GoTo 0
        End If

        'Increase the RowHeight of the ListObject on printed sheet
        If scope = HListPrintedTable Then
            Lo.DataBodyRange.EntireRow.RowHeight = 70
        End If

    Case HListHeader, HListPrintedHeader

        If (Not TypeName(formattingObject) = "Range") Then Exit Sub

        Set rng = formattingObject
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
            Next

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
Private Sub Import(ByVal inpsh As Worksheet)

    Dim targetSheet As Worksheet
    Dim inputTable As ListObject
    Dim targetTable As ListObject
    Dim inputTab As ICustomTable
    Dim targetTab As ICustomTable
    Dim designList As BetterArray
    Dim designName As String
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim rowIndex As Long

    If inpsh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Import worksheet is required"
    End If

    If inpsh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Import worksheet must expose a format table"
    End If

    If Not HasNamedRange(inpsh, "DESIGNTYPE") Then
        ThrowError ProjectError.ElementNotFound, "Range 'DESIGNTYPE' is missing from import worksheet"
    End If

    EnsureFormatReady
    Set targetSheet = Wksh()
    Set targetTable = FormatTable()
    Set inputTable = inpsh.ListObjects(1)

    'Import the design names
    targetSheet.Range("DESIGNTYPE").Value = inpsh.Range("DESIGNTYPE").Value
    DesignName = CStr(targetSheet.Range("DESIGNTYPE").Value)

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
Private Sub Export(ByVal toWkb As Workbook)

    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As ListObject
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject
    Dim exportTable As ICustomTable
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
Private Sub CopyFormatTableStyles(ByVal sourceTable As ListObject, ByVal targetTable As ListObject)
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim sourceCell As Range
    Dim targetCell As Range

    If sourceTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source format table is required for export"
    End If

    If targetTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Target format table is required for export"
    End If

    Set sourceRange = sourceTable.Range
    Set targetRange = targetTable.Range

    If sourceRange.Rows.Count <> targetRange.Rows.Count Or _
       sourceRange.Columns.Count <> targetRange.Columns.Count Then
        ThrowError ProjectError.ErrorUnexpectedState, "Export format table dimensions do not match the source"
    End If

    For rowIndex = 1 To sourceRange.Rows.Count
        For columnIndex = 1 To sourceRange.Columns.Count
            Set sourceCell = sourceRange.Cells(rowIndex, columnIndex)
            Set targetCell = targetRange.Cells(rowIndex, columnIndex)

            targetCell.Interior.Pattern = sourceCell.Interior.Pattern
            targetCell.Interior.color = sourceCell.Interior.color
            targetCell.Font.color = sourceCell.Font.color
            targetCell.Font.Bold = sourceCell.Font.Bold
            targetCell.Font.Italic = sourceCell.Font.Italic
            targetCell.Font.Size = sourceCell.Font.Size
        Next columnIndex
    Next rowIndex
End Sub

'@label:ensure-design-type-range
'@sub-title Create the DESIGNTYPE named range on the target sheet
Private Sub EnsureDesignTypeRange(ByVal sourceSheet As Worksheet, _
                                  ByVal targetSheet As Worksheet, _ 
                                  ByVal targetTable As ListObject)

    Dim sourceDesignRange As Range
    Dim targetDesignRange As Range
    Dim designValue As String
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

    If HasNamedRange(sourceSheet, "DESIGNTYPE") Then
        Set sourceDesignRange = sourceSheet.Range("DESIGNTYPE")
        designValue = CStr(sourceDesignRange.Value)
        If LenB(designValue) = 0 Then designValue = DEFAULT_DESIGN
        Set targetDesignRange = targetSheet.Cells(sourceDesignRange.Row, sourceDesignRange.Column)
    Else
        designValue = DesignName()
        If LenB(designValue) = 0 Then designValue = DEFAULT_DESIGN

        headerRow = targetTable.HeaderRowRange.Row
        headerColumn = targetTable.Range.Column
        targetColumn = headerColumn + targetTable.ListColumns.Count + 2

        Set targetDesignRange = targetSheet.Cells(headerRow, targetColumn)
    End If

    targetDesignRange.Value = designValue

    On Error Resume Next
        targetSheet.Names("DESIGNTYPE").Delete
    On Error GoTo 0
    targetSheet.Names.Add Name:="DESIGNTYPE", RefersTo:=targetDesignRange, Visible:=False
End Sub

'@label:has-named-range
'@sub-title Test whether a worksheet-scoped named range exists
'@return Boolean. True when the named range is found.
Private Function HasNamedRange(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
        Set rng = sh.Range(rngName)
    On Error GoTo 0
    HasNamedRange = Not (rng Is Nothing)
End Function

'@label:throw-error
'@sub-title Raise a typed project error
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, errorMessage
End Sub

'@section Interface
'===============================================================================
'ILLFormat delegation wiring.

Private Sub ILLFormat_ApplyFormat(formattingObject As Variant, ByVal scope As Byte, _
                                  Optional ByVal nbDecimals As Integer = 0)
    ApplyFormat formattingObject:=formattingObject, scope:=scope, nbDecimals:=nbDecimals
End Sub

Private Sub ILLFormat_Import(ByVal inpsh As Worksheet)
    Import inpsh
End Sub

Private Sub ILLFormat_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

Private Function ILLFormat_DesignValue(ByVal labelName As String, _ 
                                       Optional ByVal returnColor As Boolean = True) As Variant
    ILLFormat_DesignValue = DesignValue(labelName, returnColor)
End Function

Private Property Get ILLFormat_HasCheckings() As Boolean
    ILLFormat_HasCheckings = HasCheckings
End Property

Private Property Get ILLFormat_CheckingValues() As IChecking
    Set ILLFormat_CheckingValues = CheckingValues
End Property
