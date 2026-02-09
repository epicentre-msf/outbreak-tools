VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Writes checking results to a worksheet"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Writes checking results to a worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends BetterArray, Checking, HiddenNames

Option Explicit
Implements ICheckingOutput

'@class CheckingOutput
'@description
'Renders IChecking diagnostic entries to a worksheet with formatted rows,
'severity-based colour coding, and interactive drop-down filters for
'status and title. The class manages hidden helper columns, a title
'store for filter validation lists, and optional worksheet event
'injection so that filter selections trigger row visibility changes
'automatically.
'@version 1.0 (2026-02-09)
'@author OutbreakTools


'@section State
'===============================================================================
'@description Backing UDT, formatting structure, and constants.

Private Type TState
    targetSheet As Worksheet
    nm As String
    store As IHiddenNames
End Type

Private Type TFormatting
    fontColour As Long
    fillColour As Long
    fontSize As Long
    borderColour As Long
    fontBold As Boolean
End Type

Private this As TState


Private Const HIDDEN_TITLE_COLUMN As Long = 2 'Column B (hidden titles)
Private Const FIRST_OUTPUT_COLUMN As Long = 3 'Column C (visible data)
Private Const OUTPUT_COLUMN_COUNT As Long = 3 'Columns C:E for visible output
Private Const FIRST_ROW_OUTPUT As Long = 4
Private Const FILTER_CELL_ADDRESS As String = "C1"
Private Const FILTER_NAME As String = "CheckingOutputFilter"
Private Const TITLE_FILTER_CELL_ADDRESS As String = "E1"
Private Const TITLE_FILTER_NAME As String = "CheckingOutputTitleFilter"
Private Const TITLE_STORE_ANCHOR As String = "Z2"
Private Const TITLE_RANGE_NAME As String = "CheckingOutputTitles"
Private Const TITLE_FILTER_DEFAULT As String = "All Titles"
Private Const EVENT_MARKER_NAME As String = "CheckingOutputEventInstalled"
Private Const DEFAULT_SEPARATOR As String = "--"
Private Const ROW_MARKER_NAME As String = "CheckingOutputStartingRow"
Private Const END_TAG As String = "__end__checkings___"
Private Const VBIDE_PROC_KIND As Long = 0 'Matches vbext_pk_Proc to avoid VBIDE dependency

'@section Factory
'===============================================================================
'@description Entry point for creating CheckingOutput instances.

'@label:create
'@sub-title Create an output writer bound to a worksheet
'@details
'Instantiates a new writer and associates it with the supplied
'worksheet. Optionally assigns a title used for the end-of-checkings
'marker.
'@param sh Worksheet. Target worksheet that will receive the checking output.
'@param title Optional String. Label for the output session. Defaults to vbNullString.
'@return ICheckingOutput. Fully initialised instance.
'@throws ProjectError.ElementNotFound When sh is Nothing.
'@export
Public Function Create(ByVal sh As Worksheet, Optional ByVal title As String = vbNullString) As ICheckingOutput
    If sh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
         "Unable to create an instance of the class, please provide a valid worksheet"
    End If

    With New CheckingOutput
        Set .Wksh = sh
         .Name = title
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose internal state and the interface pointer.

'@label:wksh
'@prop-title Retrieve the worksheet backing this output writer
'@details
'Returns the worksheet previously supplied during Create or the setter.
'@return Worksheet. Reference to the target worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.targetSheet
End Property

'@label:wksh-set
'@prop-title Define the worksheet that will receive the output
'@param sh Worksheet. Target worksheet to use.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an ICheckingOutput reference.
'@return ICheckingOutput. Reference to the current object.
Public Property Get Self() As ICheckingOutput
    Set Self = Me
End Property

'@label:name
'@prop-title Name of the output session
'@details
'Returns the label assigned during creation or via the setter.
'@return String. Output session label.
Public Property Get Name() As String
    Name = this.nm
End Property

'@label:name-set
'@prop-title Assign the output session name
'@param nm String. Label to assign.
Public Property Let Name(ByVal nm As String)
    this.nm = nm
End Property

'@section FormattingHelpers
'===============================================================================
'@description Colour resolution, text decoration, and border styling.

'@label:resolve-formatting
'@sub-title Resolve formatting attributes for a given level and colour name
'@details
'Maps a tag string (colour name or heading level) to font, fill, and
'border properties returned in a TFormatting structure.
'@param tag Optional String. Colour or level type token (e.g. "red", "level 1").
'@return TFormatting. Structure with font, fill, and border properties.
Private Function ResolveFormatting(Optional ByVal tag As String) As TFormatting
    Dim formatInfo As TFormatting


    Select Case LCase$(tag)
        Case "level 1"
            formatInfo.fontBold = True
            formatInfo.fontSize = 20
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64)          ' Dark navy border
        Case "level 2"
            formatInfo.fontBold = True
            formatInfo.fontSize = 14
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64)          ' Dark navy border
        Case "purple"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(112, 48, 160)
            formatInfo.fillColour = RGB(244, 236, 255)
            formatInfo.borderColour = RGB(203, 167, 226)
        Case "red"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(192, 0, 0)
            formatInfo.fillColour = RGB(255, 235, 238)
            formatInfo.borderColour = RGB(224, 102, 102)
        Case "orange"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(167, 106, 0)
            formatInfo.fillColour = RGB(255, 240, 220)
            formatInfo.borderColour = RGB(236, 184, 120)
        Case "grey"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(96, 97, 100)
            formatInfo.fillColour = RGB(242, 243, 245)
            formatInfo.borderColour = RGB(205, 206, 210)
        Case "green"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(0, 120, 50)
            formatInfo.fillColour = RGB(220, 245, 226)
            formatInfo.borderColour = RGB(120, 200, 140)
        Case "label"
            'Default formatting
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(0, 0, 128)           ' Greenish font
            formatInfo.fillColour = vbWhite                  ' Light green fill
            formatInfo.borderColour = RGB(0, 0, 64)
        Case Else
            'Default formatting
            formatInfo.fontSize = 9
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64)
    End Select

    ResolveFormatting = formatInfo
End Function


'@label:apply-text-decorations
'@sub-title Apply typography styling based on formatting
'@details
'Decorates the target range with font size, colour, fill, and bold
'emphasis from the supplied formatting structure.
'@param formatting TFormatting. Structure containing font and fill attributes.
'@param target Range. Range receiving the formatting.
Private Sub ApplyTextDecorations(formatting As TFormatting, ByVal target As Range)

    With target
        .Font.Size = formatting.fontSize
        .Font.Color = formatting.fontColour
        .Interior.Color = formatting.fillColour
        .Font.Bold = formatting.fontBold
    End With
End Sub


'@label:apply-borders
'@sub-title Apply border styling to a range
'@details
'Applies a dashed bottom border to the target range block.
'@param targetRange Range. Range receiving the border formatting.
'@param borderColour Long. Colour code for the border.
Private Sub ApplyBorders(ByVal targetRange As Range, ByVal borderColour As Long)
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .Weight = xlThin
        .Color = borderColour
    End With
End Sub


'@label:write-row
'@sub-title Format and write a single row of checking output
'@details
'Splits the text value by the separator and writes the fragments
'across the output columns. Applies formatting based on the format
'tag and records the parent title in the hidden column for filtering.
'@param rowIndex Long. Row where to write.
'@param textValue String. Content to write, split by sep.
'@param formatTag Optional String. Tag for formatting resolution.
'@param sep Optional String. Split separator. Defaults to "--".
'@param parentTitle Optional String. Title written to the hidden column.
Private Sub WriteRow(ByVal rowIndex As Long, ByVal textValue As String, _
                     Optional ByVal formatTag As String, _
                     Optional ByVal sep As String = DEFAULT_SEPARATOR, _
                     Optional ByVal parentTitle As String = vbNullString)
    Dim target As Range
    Dim splitValues() As String
    Dim formatting As TFormatting
    Dim sh As Worksheet
    Dim hiddenCell As Range

    Set sh = Wksh()

    Set target = sh.Range(sh.Cells(rowIndex, FIRST_OUTPUT_COLUMN), _
                          sh.Cells(rowIndex, FIRST_OUTPUT_COLUMN + OUTPUT_COLUMN_COUNT - 1))

    Set hiddenCell = sh.Cells(rowIndex, HIDDEN_TITLE_COLUMN)
    hiddenCell.NumberFormat = "@"
    hiddenCell.Font.Color = vbWhite

    If LenB(parentTitle) = 0 Then
        hiddenCell.value = vbNullString
    Else
        hiddenCell.value = parentTitle
    End If

    splitValues = Split(textValue, sep)
    target.Cells(1, 1).value = ValueOrEmpty(splitValues, 0)
    target.Cells(1, 2).value = ValueOrEmpty(splitValues, 1)
    target.Cells(1, 3).value = ValueOrEmpty(splitValues, 2)

    formatting = ResolveFormatting(formatTag)

    ApplyTextDecorations formatting, target
    target.EntireColumn.AutoFit
    ApplyBorders target, formatting.borderColour
End Sub

'@label:value-or-empty
'@sub-title Safely retrieve an element from a split array
'@details
'Returns the element at the specified index when it exists, or
'vbNullString when the index exceeds the array bounds.
'@param items String(). Split array.
'@param index Long. Zero-based index to retrieve.
'@return String. Element value or vbNullString.
Private Function ValueOrEmpty(items() As String, ByVal index As Long) As String
    If index > UBound(items) Then
        ValueOrEmpty = vbNullString
    Else
        ValueOrEmpty = items(index)
    End If
End Function

'@section CheckingWriters
'===============================================================================
'@description Outputs headings and detail entries for individual checking bundles.

'@label:write-checking
'@sub-title Write a single IChecking bundle to the worksheet
'@details
'Outputs the title heading (once per unique title), optional subtitle,
'and all keyed entries with their severity colour coding. Advances the
'current row counter and tracks written titles to avoid duplicates.
'@param checking IChecking. Bundle to render.
'@param currentRow Long. ByRef row counter advanced during writing.
'@param writtenTitles Collection. ByRef collection tracking emitted titles.
'@param sep Optional String. Separator for text fragments. Defaults to "--".
Private Sub WriteChecking(ByVal checking As IChecking, _
                          ByRef currentRow As Long, _
                          ByRef writtenTitles As Collection, _
                          Optional ByVal sep As String = DEFAULT_SEPARATOR)
'Outputs headings and detail entries for an individual checking bundle.

    Dim title As String
    Dim subtitle As String
    Dim keys As BetterArray
    Dim keyIndex As Long
    Dim keyName As String
    Dim keyColour As String
    Dim keyLabel As String
    Dim keyType As String
    Dim fullLabel As String

    title = checking.Heading
    subtitle = checking.Heading(subtitle:=True)

    If Not TitleAlreadyWritten(title, writtenTitles) Then
        WriteRow rowIndex:=currentRow, textValue:=title, formatTag:="level 1", sep:=sep, parentTitle:=title
        If LenB(Trim$(title)) > 0 Then AppendTitleToStore title
        writtenTitles.Add Item:=title, key:=title
        currentRow = currentRow + 2
    End If

    If subtitle <> vbNullString Then
        currentRow = currentRow + 1
        WriteRow rowIndex:=currentRow, textValue:=subtitle, formatTag:="level 2", sep:=sep, parentTitle:=title
        currentRow = currentRow + 2
    End If

    Set keys = checking.ListOfKeys
    For keyIndex = keys.lowerBound To keys.upperBound
        keyName = CStr(keys.Item(keyIndex))
        keyColour = checking.ValueOf(keyName, checkingColor)
        keyLabel = checking.ValueOf(keyName, checkingLabel)
        keyType = checking.ValueOf(keyName, checkingType)
        fullLabel = keyType & sep & keyLabel
        WriteRow rowIndex:=currentRow, textValue:=fullLabel, formatTag:=keyColour, sep:=sep, parentTitle:=title
        currentRow = currentRow + 1
    Next keyIndex

    currentRow = currentRow + 2
End Sub


'@label:title-already-written
'@sub-title Test whether a title has already been emitted
'@details
'Checks the writtenTitles collection for the specified title.
'Uses On Error Resume Next to handle missing keys.
'@param title String. Title to look for.
'@param writtenTitles Collection. Collection of previously emitted titles.
'@return Boolean. True when the title was already written.
Private Function TitleAlreadyWritten(ByVal title As String, ByRef writtenTitles As Collection) As Boolean
    On Error Resume Next
        TitleAlreadyWritten = Not IsEmpty(writtenTitles.Item(title))
    On Error GoTo 0
End Function

'@section TitleStoreManagement
'===============================================================================
'@description Manages the hidden title store used for the title filter validation list.

'@label:ensure-title-store
'@sub-title Initialise or refresh the hidden title store
'@details
'When resetStore is True, clears all stored titles and re-creates
'the anchor cell. Otherwise ensures the anchor exists and the named
'range is defined.
'@param resetStore Optional Boolean. When True, clears the store. Defaults to False.
Private Sub EnsureTitleStore(Optional ByVal resetStore As Boolean = False)
    Dim sh As Worksheet

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    If resetStore Then
        ResetTitleStore sh
        Exit Sub
    End If

    EnsureTitleStoreAnchor sh

    If Not WorksheetHasNamedRange(sh, TITLE_RANGE_NAME) Then
        RefreshTitleNamedRange sh
    End If

End Sub

'@label:reset-title-store
'@sub-title Clear all stored titles and re-create the anchor
'@param sh Worksheet. Target worksheet.
Private Sub ResetTitleStore(ByVal sh As Worksheet)
    Dim anchor As Range
    Dim clearRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    Set clearRange = sh.Range(anchor, sh.Cells(sh.Rows.Count, anchor.Column))
    clearRange.ClearContents

    anchor.NumberFormat = "@"
    anchor.value = TITLE_FILTER_DEFAULT
    anchor.Font.Color = vbWhite

    On Error Resume Next
        sh.Names(TITLE_RANGE_NAME).Delete
    On Error GoTo 0

    RefreshTitleNamedRange sh
End Sub

'@label:ensure-title-store-anchor
'@sub-title Ensure the title store anchor cell is initialised
'@param sh Worksheet. Target worksheet.
Private Sub EnsureTitleStoreAnchor(ByVal sh As Worksheet)
    Dim anchor As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    anchor.NumberFormat = "@"
    If LenB(Trim$(CStr(anchor.value))) = 0 Then
        anchor.value = TITLE_FILTER_DEFAULT
    End If

    anchor.Font.Color = vbWhite
End Sub

'@label:append-title-to-store
'@sub-title Add a title to the hidden store if not already present
'@param titleValue String. Title text to store.
Private Sub AppendTitleToStore(ByVal titleValue As String)
    Dim sh As Worksheet
    Dim anchor As Range
    Dim target As Range
    Dim nextRow As Long

    titleValue = Trim$(titleValue)
    If LenB(titleValue) = 0 Then Exit Sub

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    EnsureTitleStoreAnchor sh

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    If TitleStoreContains(sh, titleValue) Then Exit Sub

    nextRow = NextTitleStoreRow(sh, anchor)
    Set target = sh.Cells(nextRow, anchor.Column)

    target.value = titleValue
    target.NumberFormat = "@"
    target.Font.Color = vbWhite

    RefreshTitleNamedRange sh
End Sub

'@label:title-store-contains
'@sub-title Check whether the title store already includes a value
'@param sh Worksheet. Target worksheet.
'@param titleValue String. Title text to search for.
'@return Boolean. True when the title is already stored.
Private Function TitleStoreContains(ByVal sh As Worksheet, ByVal titleValue As String) As Boolean
    Dim anchor As Range
    Dim lastRow As Long
    Dim titleRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then Exit Function

    Set titleRange = sh.Range(anchor, sh.Cells(lastRow, anchor.Column))

    On Error Resume Next
        TitleStoreContains = WorksheetFunction.CountIf(titleRange, titleValue) > 0
    On Error GoTo 0
End Function

'@label:refresh-title-named-range
'@sub-title Update the named range covering all stored titles
'@param sh Worksheet. Target worksheet.
Private Sub RefreshTitleNamedRange(ByVal sh As Worksheet)
    Dim anchor As Range
    Dim lastRow As Long
    Dim titleRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then lastRow = anchor.Row

    Set titleRange = sh.Range(anchor, sh.Cells(lastRow, anchor.Column))

    On Error Resume Next
        sh.Names(TITLE_RANGE_NAME).Delete
    On Error GoTo 0

    sh.Names.Add Name:=TITLE_RANGE_NAME, RefersTo:=titleRange
End Sub

'@label:next-title-store-row
'@sub-title Determine the next available row in the title store
'@param sh Worksheet. Target worksheet.
'@param anchor Range. Anchor cell of the title store.
'@return Long. Next available row index.
Private Function NextTitleStoreRow(ByVal sh As Worksheet, ByVal anchor As Range) As Long
    Dim lastRow As Long

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then
        NextTitleStoreRow = anchor.Row + 1
    Else
        NextTitleStoreRow = lastRow + 1
    End If
End Function

'@label:worksheet-has-named-range
'@sub-title Check whether a named range exists on a worksheet
'@param targetSheet Worksheet. Worksheet to inspect.
'@param rangeName String. Name to look for.
'@return Boolean. True when the named range exists.
Private Function WorksheetHasNamedRange(ByVal targetSheet As Worksheet, ByVal rangeName As String) As Boolean
    Dim nm As Name

    On Error Resume Next
        Set nm = targetSheet.Names(rangeName)
        WorksheetHasNamedRange = Err.Number = 0 And Not nm Is Nothing
    On Error GoTo 0
End Function

'@label:worksheet-name-for-formula
'@sub-title Escape the worksheet name for use in formulas
'@param targetSheet Worksheet. Worksheet whose name to escape.
'@return String. Escaped worksheet name.
Private Function WorksheetNameForFormula(ByVal targetSheet As Worksheet) As String
    WorksheetNameForFormula = Replace(targetSheet.Name, "'", "''")
End Function

'@label:apply-hidden-columns-formatting
'@sub-title Mask hidden helper columns so metadata remains invisible
'@details
'Ensures the parent-title column and backing title store column
'keep a white font and hidden presentation.
'@param sh Worksheet. Target worksheet to format.
Private Sub ApplyHiddenColumnsFormatting(ByVal sh As Worksheet)
    Dim storeColumn As Long

    If sh Is Nothing Then Exit Sub

    storeColumn = sh.Range(TITLE_STORE_ANCHOR).Column

    With sh.Columns(HIDDEN_TITLE_COLUMN)
        .NumberFormat = "@"
        .Font.Color = vbWhite
        .Interior.Color = vbWhite
    End With

    With sh.Columns(storeColumn)
        .NumberFormat = "@"
        .Font.Color = vbWhite
        .Interior.Color = vbWhite
        .Hidden = True
    End With
End Sub

'@section WorksheetPreparation
'===============================================================================
'@description Resets and initialises the worksheet before rendering.

'@label:initialise-worksheet
'@sub-title Reset the worksheet before re-rendering checking output
'@details
'On first render, clears all cells, sets the Consolas font, creates
'the title store, installs filter dropdowns, removes gridlines, and
'records the starting row marker. On subsequent renders, resumes from
'the last written row.
'@return Long. Starting row for the next output block.
Private Function InitialiseWorksheet() As Long
    Dim sh As Worksheet
    Dim startingRow As Long
    Dim view As WorksheetView
    Dim store As IHiddenNames

    Set sh = Wksh()
    sh.Rows.Hidden = False
    Set store = HiddenStore()

    If Not store.HasName(ROW_MARKER_NAME) Then
        sh.Cells.Clear
        sh.Cells.Font.Name = "Consolas"
        sh.rows(1).Font.Bold = True
        sh.rows(1).Font.Size = 14

        EnsureTitleStore True
        ApplyHiddenColumnsFormatting sh
        EnsureFilterDropdown
        startingRow = FIRST_ROW_OUTPUT
        store.EnsureName ROW_MARKER_NAME, FIRST_ROW_OUTPUT, HiddenNameTypeLong

        'Remove gridlines of the worksheet
        For Each view In sh.Parent.Windows(1).SheetViews
            If view.sheet.Name = sh.Name Then
                view.DisplayGridlines = False
                view.DisplayZeros = False
                Exit For
            End If
        Next view
    Else
        startingRow = sh.Cells(sh.Rows.Count, FIRST_OUTPUT_COLUMN).End(xlUp).Row
        If startingRow < FIRST_ROW_OUTPUT Then
            startingRow = FIRST_ROW_OUTPUT
            EnsureTitleStore True
        Else
            EnsureTitleStore False
            sh.Rows(startingRow).ClearContents
        End If
        ApplyHiddenColumnsFormatting sh
        EnsureFilterDropdown

    End If

    InitialiseWorksheet = startingRow
End Function



'@section Filtering
'===============================================================================
'@description Drop-down filter management and row-visibility algorithms.

'@label:ensure-filter-dropdown
'@sub-title Install and style the filter drop-downs on the worksheet header
'@details
'Creates validation drop-downs for both status (severity) and title
'filtering. Ensures the title store named range is available for the
'title validation list.
Private Sub EnsureFilterDropdown()
    Dim sh As Worksheet
    Dim statusCell As Range
    Dim titleCell As Range
    Dim statusLabel As Range
    Dim titleLabel As Range

    Set sh = Wksh()

    EnsureTitleStore False

    Set statusLabel = sh.Cells(1, HIDDEN_TITLE_COLUMN)
    statusLabel.value = "Status:"
    ApplyTextDecorations ResolveFormatting("label"), statusLabel

    Set statusCell = sh.Range(FILTER_CELL_ADDRESS)
    ApplyTextDecorations ResolveFormatting("dropdown"), statusCell

    On Error Resume Next
        sh.Names(FILTER_NAME).Delete
    On Error GoTo 0
    sh.Names.Add Name:=FILTER_NAME, RefersTo:=statusCell

    With statusCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              Formula1:="All,Errors,Warnings,Notes,Infos,Successes,Without Successes"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    If LenB(Trim$(CStr(statusCell.value))) = 0 Then statusCell.value = "All"

    Set titleLabel = sh.Cells(1, sh.Range(TITLE_FILTER_CELL_ADDRESS).Column - 1)
    titleLabel.value = "Title:"
    ApplyTextDecorations ResolveFormatting("label"), titleLabel

    Set titleCell = sh.Range(TITLE_FILTER_CELL_ADDRESS)
    ApplyTextDecorations ResolveFormatting("dropdown"), titleCell

    On Error Resume Next
        sh.Names(TITLE_FILTER_NAME).Delete
    On Error GoTo 0

    sh.Names.Add Name:=TITLE_FILTER_NAME, RefersTo:=titleCell

    RefreshTitleNamedRange sh

    With titleCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              Formula1:="= " & TITLE_RANGE_NAME
        .IgnoreBlank = True
        .InCellDropdown = True
    End With


    If LenB(Trim$(CStr(titleCell.value))) = 0 Then titleCell.value = TITLE_FILTER_DEFAULT
End Sub

'@label:filter-worksheet
'@sub-title Apply filtering logic using the backing worksheet
'@details
'Public entry point that allows consumers to reapply the worksheet
'filtering logic without relying on events. Delegates to
'ApplyWorksheetFilter.
'@param statusSelection Optional String. Manual filter selection overriding the drop-down value.
'@param titleSelection Optional String. Title filter overriding the worksheet selection.
'@export
Public Sub FilterWorksheet(Optional ByVal statusSelection As String = vbNullString, _
                          Optional ByVal titleSelection As String = vbNullString)
    If Wksh() Is Nothing Then Exit Sub
    ApplyWorksheetFilter Wksh, statusSelection, titleSelection
End Sub

'@label:apply-worksheet-filter
'@sub-title Internal helper that performs row-visibility filtering
'@details
'Centralises the filtering algorithm so both the public API and
'worksheet events reuse the same logic. Hides rows that do not match
'the selected severity and/or title criteria.
'@param targetSheet Worksheet. Worksheet hosting the checking output data.
'@param statusSelection Optional String. Desired severity filter; blank uses the worksheet cell.
'@param titleSelection Optional String. Title token overriding the drop-down when provided.
Private Sub ApplyWorksheetFilter(ByVal targetSheet As Worksheet, _
                                 Optional ByVal statusSelection As String = vbNullString, _
                                 Optional ByVal titleSelection As String = vbNullString)
    Dim statusValue As String
    Dim titleValue As String
    Dim matchMode As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowType As String
    Dim hasLabel As Boolean
    Dim rowTitle As String
    Dim filterByTitle As Boolean
    Dim explicitStatusOverride As Boolean

    If targetSheet Is Nothing Then Exit Sub

    statusValue = Trim$(statusSelection)
    explicitStatusOverride = LenB(statusValue) > 0
    If LenB(statusValue) = 0 Then
        statusValue = Trim$(CStr(targetSheet.Range(FILTER_CELL_ADDRESS).value))
    End If

    titleValue = Trim$(titleSelection)
    If LenB(titleValue) = 0 Then
        titleValue = Trim$(CStr(targetSheet.Range(TITLE_FILTER_CELL_ADDRESS).value))
    End If

    filterByTitle = LenB(titleValue) > 0 And _
                    StrComp(titleValue, TITLE_FILTER_DEFAULT, vbTextCompare) <> 0

    Select Case UCase$(statusValue)
        Case "", "ALL"
            matchMode = "ALL"
        Case "ERROR", "ERRORS"
            matchMode = "ERROR"
        Case "WARNING", "WARNINGS"
            matchMode = "WARNING"
        Case "NOTE", "NOTES"
            matchMode = "NOTE"
        Case "INFO", "INFOS"
            matchMode = "INFO"
        Case "SUCCESS", "SUCCESSES"
            matchMode = "SUCCESS"
        Case "WITHOUT SUCCESS", "WITHOUT SUCCESSES"
            matchMode = "WITHOUT_SUCCESS"
        Case Else
            matchMode = "ALL"
    End Select

    If filterByTitle And Not explicitStatusOverride Then
        'Grant the title filter precedence when no status override is explicitly supplied
        matchMode = "ALL"
    End If

    lastRow = targetSheet.Cells(targetSheet.rows.Count, FIRST_OUTPUT_COLUMN).End(xlUp).row
    If lastRow < FIRST_ROW_OUTPUT Then Exit Sub

    For rowIndex = FIRST_ROW_OUTPUT To lastRow
        targetSheet.rows(rowIndex).Hidden = False
        rowTitle = Trim$(CStr(targetSheet.Cells(rowIndex, HIDDEN_TITLE_COLUMN).value))

        If filterByTitle Then
            If StrComp(rowTitle, titleValue, vbTextCompare) <> 0 Then
                targetSheet.rows(rowIndex).Hidden = True
                GoTo ContinueLoop
            End If
        End If

        hasLabel = LenB(CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN + 1).value)) > 0

        If hasLabel Then
            rowType = NormaliseTypeLabel(CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN).value))
            Select Case matchMode
                Case "ALL"
                    ' no filtering
                Case "SUCCESS"
                    If StrComp(rowType, "Success", vbTextCompare) <> 0 Then targetSheet.rows(rowIndex).Hidden = True
                Case "WITHOUT_SUCCESS"
                    If StrComp(rowType, "Success", vbTextCompare) = 0 Then targetSheet.rows(rowIndex).Hidden = True
                Case Else
                    If StrComp(rowType, matchMode, vbTextCompare) <> 0 Then targetSheet.rows(rowIndex).Hidden = True
            End Select
        End If

ContinueLoop:
    Next rowIndex
End Sub

'@section WorksheetEvents
'===============================================================================
'@description Injection and generation of the Worksheet_Change handler for interactive filtering.

'@label:ensure-worksheet-change-handler
'@sub-title Inject the Worksheet_Change handler into the target module
'@details
'Injects the filtering VBA code into the worksheet's code module so
'filter drop-down changes trigger row visibility updates automatically.
'Skips injection when the event marker hidden name is already present.
'@throws ProjectError.SomethingWentWrong When code injection fails.
'@export
Private Sub EnsureWorksheetChangeHandler()
    Dim codeModule As Object
    Dim injectionCode As String
    Dim moduleLines As Long
    Dim store As IHiddenNames

    On Error GoTo Handler

    If Wksh() Is Nothing Then Exit Sub
    Set store = HiddenStore()
    If store.HasName(EVENT_MARKER_NAME) Then Exit Sub

    Set codeModule = Wksh.Parent.VBProject.VBComponents(Wksh.CodeName).CodeModule

    RemoveManagedWorksheetProcedures codeModule

    injectionCode = BuildWorksheetChangeHandlerCode()
    moduleLines = codeModule.CountOfLines

    If moduleLines > 0 Then
        codeModule.InsertLines moduleLines + 1, vbNewLine & injectionCode
    Else
        codeModule.InsertLines 1, injectionCode
    End If

    store.EnsureName EVENT_MARKER_NAME, True, HiddenNameTypeBoolean
    Exit Sub

Handler:
    ThrowError ProjectError.SomethingWentWrong, _
        "Unable to install worksheet change handler: " & Err.Description
End Sub

'@label:build-worksheet-change-handler-code
'@sub-title Compose the VBA code for the worksheet filtering logic
'@details
'Generates the complete VBA source code that is injected into the
'worksheet module. Includes StripIcons, ResolveMatchType,
'ShouldFilterByTitle, RestoreState, Worksheet_Change, and
'FilterCheckingOutputRows procedures.
'@return String. VBA source code for the handler.
Private Function BuildWorksheetChangeHandlerCode() As String
    Dim handler As String
    Dim dq As String

    dq = """"

    handler = handler & "Private Function StripIcons(ByVal txt As String) As String" & vbNewLine
    handler = handler & "    Dim stripped As String" & vbNewLine
    handler = handler & "    stripped = Replace(txt, ChrW(10060), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(9888), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(8505), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(9998), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(10004), """")" & vbNewLine
    handler = handler & "    StripIcons = Trim$(stripped)" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Function ResolveMatchType(ByVal candidate As String) As String" & vbNewLine
    handler = handler & "    Select Case UCase$(Trim$(candidate))" & vbNewLine
    handler = handler & "        Case """", ""ALL""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ALL""" & vbNewLine
    handler = handler & "        Case ""ERROR"", ""ERRORS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ERROR""" & vbNewLine
    handler = handler & "        Case ""WARNING"", ""WARNINGS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""WARNING""" & vbNewLine
    handler = handler & "        Case ""NOTE"", ""NOTES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""NOTE""" & vbNewLine
    handler = handler & "        Case ""INFO"", ""INFOS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""INFO""" & vbNewLine
    handler = handler & "        Case ""SUCCESS"", ""SUCCESSES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""SUCCESS""" & vbNewLine
    handler = handler & "        Case ""WITHOUT SUCCESS"", ""WITHOUT SUCCESSES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""WITHOUT_SUCCESS""" & vbNewLine
    handler = handler & "        Case Else" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ALL""" & vbNewLine
    handler = handler & "    End Select" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Function ShouldFilterByTitle(ByVal candidate As String) As Boolean" & vbNewLine
    handler = handler & "    Dim trimmedTitle As String" & vbNewLine
    handler = handler & "    trimmedTitle = Trim$(candidate)" & vbNewLine
    handler = handler & "    ShouldFilterByTitle = LenB(trimmedTitle) > 0 And StrComp(trimmedTitle, " & dq & TITLE_FILTER_DEFAULT & dq & ", vbTextCompare) <> 0" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Sub RestoreState(ByVal previousEvents As Boolean, ByVal previousUpdating As Boolean, ByVal previousCalculation As XlCalculation)" & vbNewLine
    handler = handler & "    Application.EnableEvents = previousEvents" & vbNewLine
    handler = handler & "    Application.ScreenUpdating = previousUpdating" & vbNewLine
    handler = handler & "    Application.Calculation = previousCalculation" & vbNewLine
    handler = handler & "End Sub" & vbNewLine & vbNewLine

    handler = handler & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine
    handler = handler & "    On Error GoTo Handler" & vbNewLine
    handler = handler & "    Dim statusCell As Range" & vbNewLine
    handler = handler & "    Dim titleCell As Range" & vbNewLine
    handler = handler & "    Dim statusChanged As Boolean" & vbNewLine
    handler = handler & "    Dim titleChanged As Boolean" & vbNewLine
    handler = handler & "    Dim previousEvents As Boolean" & vbNewLine
    handler = handler & "    Dim previousUpdating As Boolean" & vbNewLine
    handler = handler & "    Dim previousCalculation As XlCalculation" & vbNewLine
    handler = handler & "    Dim managedState As Boolean" & vbNewLine
    handler = handler & "    Dim errorNumber As Long" & vbNewLine
    handler = handler & "    Dim errorDescription As String" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    If Target Is Nothing Then Exit Sub" & vbNewLine
    handler = handler & "    If Target.CountLarge > 1 Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    Set statusCell = Me.Range(" & dq & FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handler = handler & "    Set titleCell = Me.Range(" & dq & TITLE_FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    statusChanged = Not Intersect(Target, statusCell) Is Nothing" & vbNewLine
    handler = handler & "    titleChanged = Not Intersect(Target, titleCell) Is Nothing" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    If Not statusChanged And Not titleChanged Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    previousEvents = Application.EnableEvents" & vbNewLine
    handler = handler & "    previousUpdating = Application.ScreenUpdating" & vbNewLine
    handler = handler & "    previousCalculation = Application.Calculation" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    Application.EnableEvents = False" & vbNewLine
    handler = handler & "    Application.ScreenUpdating = False" & vbNewLine
    handler = handler & "    Application.Calculation = xlCalculationManual" & vbNewLine
    handler = handler & "    managedState = True" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    FilterCheckingOutputRows CStr(statusCell.Value), CStr(titleCell.Value)" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "ExitHandler:" & vbNewLine
    handler = handler & "    If managedState Then" & vbNewLine
    handler = handler & "        RestoreState previousEvents, previousUpdating, previousCalculation" & vbNewLine
    handler = handler & "    End If" & vbNewLine
    handler = handler & "    If errorNumber <> 0 Then" & vbNewLine
    handler = handler & "        Err.Raise errorNumber, " & dq & "CheckingOutputWorksheet" & dq & ", errorDescription" & vbNewLine
    handler = handler & "    End If" & vbNewLine
    handler = handler & "    Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "Handler:" & vbNewLine
    handler = handler & "    errorNumber = Err.Number" & vbNewLine
    handler = handler & "    errorDescription = Err.Description" & vbNewLine
    handler = handler & "    Resume ExitHandler" & vbNewLine
    handler = handler & "End Sub" & vbNewLine & vbNewLine

    handler = handler & "Private Sub FilterCheckingOutputRows(ByVal statusValue As String, ByVal titleValue As String)" & vbNewLine
    handler = handler & "    Dim matchType As String" & vbNewLine
    handler = handler & "    Dim applyTitleFilter As Boolean" & vbNewLine
    handler = handler & "    Dim lastRow As Long" & vbNewLine
    handler = handler & "    Dim rowIndex As Long" & vbNewLine
    handler = handler & "    Dim rowType As String" & vbNewLine
    handler = handler & "    Dim rowTitle As String" & vbNewLine
    handler = handler & "    Dim hasLabel As Boolean" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    matchType = ResolveMatchType(statusValue)" & vbNewLine
    handler = handler & "    applyTitleFilter = ShouldFilterByTitle(titleValue)" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    lastRow = Me.Cells(Me.Rows.Count, " & FIRST_OUTPUT_COLUMN & ").End(xlUp).Row" & vbNewLine
    handler = handler & "    If lastRow < " & FIRST_ROW_OUTPUT & " Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    For rowIndex = " & FIRST_ROW_OUTPUT & " To lastRow" & vbNewLine
    handler = handler & "        Me.Rows(rowIndex).Hidden = False" & vbNewLine
    handler = handler & "        rowTitle = Trim$(CStr(Me.Cells(rowIndex, " & HIDDEN_TITLE_COLUMN & ").Value))" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        If applyTitleFilter Then" & vbNewLine
    handler = handler & "            If (StrComp(rowTitle, Trim$(titleValue), vbTextCompare) <> 0) And (rowTitle <> " & dq & END_TAG & dq & ") Then" & vbNewLine
    handler = handler & "                Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                GoTo ContinueRow" & vbNewLine
    handler = handler & "            End If" & vbNewLine
    handler = handler & "        End If" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        hasLabel = LenB(CStr(Me.Cells(rowIndex, " & (FIRST_OUTPUT_COLUMN + 1) & ").Value)) > 0" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        If hasLabel Then" & vbNewLine
    handler = handler & "            rowType = StripIcons(CStr(Me.Cells(rowIndex, " & FIRST_OUTPUT_COLUMN & ").Value))" & vbNewLine
    handler = handler & "            Select Case matchType" & vbNewLine
    handler = handler & "                Case ""ALL""" & vbNewLine
    handler = handler & "                    'No filtering required" & vbNewLine
    handler = handler & "                Case ""SUCCESS""" & vbNewLine
    handler = handler & "                    If StrComp(rowType, ""Success"", vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                Case ""WITHOUT_SUCCESS""" & vbNewLine
    handler = handler & "                    If StrComp(rowType, ""Success"", vbTextCompare) = 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                Case Else" & vbNewLine
    handler = handler & "                    If StrComp(rowType, matchType, vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "            End Select" & vbNewLine
    handler = handler & "        End If" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "ContinueRow:" & vbNewLine
    handler = handler & "    Next rowIndex" & vbNewLine
    handler = handler & "End Sub" & vbNewLine

    BuildWorksheetChangeHandlerCode = handler
End Function

'@section WorksheetEventHelpers
'===============================================================================
'@description Utilities for managing injected worksheet procedures.

'@label:remove-managed-worksheet-procedures
'@sub-title Strip previously injected worksheet routines
'@details
'Removes any instance of the managed procedures so a single fresh
'handler can be written without duplication.
'@param codeModule Object. Worksheet code module reference.
Private Sub RemoveManagedWorksheetProcedures(ByVal codeModule As Object)
    DeleteProcedureIfExists codeModule, "Worksheet_Change"
    DeleteProcedureIfExists codeModule, "StripIcons"
    DeleteProcedureIfExists codeModule, "ResolveMatchType"
    DeleteProcedureIfExists codeModule, "ShouldFilterByTitle"
    DeleteProcedureIfExists codeModule, "RestoreState"
    DeleteProcedureIfExists codeModule, "FilterCheckingOutputRows"
End Sub

'@label:delete-procedure-if-exists
'@sub-title Safely remove a procedure from a worksheet module
'@details
'Deletes the target procedure when present. Silently exits when the
'procedure is not found.
'@param codeModule Object. Worksheet code module reference.
'@param procedureName String. Name of the procedure to delete.
Private Sub DeleteProcedureIfExists(ByVal codeModule As Object, ByVal procedureName As String)
    Dim startLine As Long
    Dim lineCount As Long

    On Error Resume Next
        startLine = codeModule.ProcStartLine(procedureName, VBIDE_PROC_KIND)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        lineCount = codeModule.ProcCountLines(procedureName, VBIDE_PROC_KIND)
        codeModule.DeleteLines startLine, lineCount
    On Error GoTo 0
End Sub

'@section Helpers
'===============================================================================
'@description Private utility methods for validation, normalisation, and storage.

'@label:validate-check-table
'@sub-title Verify that the BetterArray contains only IChecking entries
'@details
'Iterates the array and raises an error for any item that does not
'implement IChecking.
'@param checkTable BetterArray. Table of checking entries to validate.
'@throws ProjectError.InvalidArgument When an item does not implement IChecking.
Private Sub ValidateCheckTable(ByVal checkTable As BetterArray)
    Dim index As Long
    For index = checkTable.lowerBound To checkTable.upperBound
        If Not TypeOf checkTable.Item(index) Is IChecking Then
           ThrowError ProjectError.InvalidArgument, _
            "Items in PrintOutput must implement IChecking. Item " & CStr(index) & " does not."
        End If
    Next index
End Sub


'@label:normalise-type-label
'@sub-title Strip icon characters from a checking type label
'@details
'Removes Unicode icon characters used for visual decoration to obtain
'the plain-text type label for filtering comparisons.
'@param typeLabel String. Label to clean.
'@return String. Cleaned label.
Private Function NormaliseTypeLabel(ByVal typeLabel As String) As String
    Dim cleaned As String

    cleaned = typeLabel
    cleaned = Replace(cleaned, ChrW(10060), vbNullString)
    cleaned = Replace(cleaned, ChrW(9888), vbNullString)
    cleaned = Replace(cleaned, ChrW(8505), vbNullString)
    cleaned = Replace(cleaned, ChrW(9998), vbNullString)
    cleaned = Replace(cleaned, ChrW(10004), vbNullString)

    NormaliseTypeLabel = Trim$(cleaned)
End Function

'@label:custom-property-exists
'@sub-title Check whether a custom property exists in the hidden store
'@param propName String. Property name to check.
'@return Boolean. True when the property exists.
Private Function CustomPropertyExists(ByVal propName As String) As Boolean
    Dim store As IHiddenNames
    Set store = HiddenStore()
    CustomPropertyExists = store.HasName(propName)
End Function

'@label:store-custom-property
'@sub-title Store a custom property in the hidden name store
'@param propName String. Property name to store.
'@param propValue String. Value to persist.
Private Sub StoreCustomProperty(propName As String, propValue As String)
    Dim store As IHiddenNames
    Set store = HiddenStore()
    store.EnsureName propName, propValue, HiddenNameTypeString
    store.SetValue propName, propValue
End Sub

'@label:hidden-store
'@sub-title Lazily create and return the IHiddenNames store
'@details
'Creates the HiddenNames wrapper bound to the target worksheet on
'first access. Raises an error when no worksheet is assigned.
'@return IHiddenNames. Hidden name store.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
Private Function HiddenStore() As IHiddenNames
    If this.store Is Nothing Then
        If Wksh() Is Nothing Then
            ThrowError ProjectError.ObjectNotInitialized, _
             "CheckingOutput requires a worksheet before accessing hidden storage."
        End If
        Set this.store = HiddenNames.Create(Wksh())
    End If
    Set HiddenStore = this.store
End Function

'@label:module-contains-procedure
'@sub-title Determine whether the module text declares a target procedure
'@details
'Searches for Sub or Function signatures matching the procedure name.
'@param moduleText String. VBA module contents.
'@param procedureName String. Name of the procedure to search for.
'@return Boolean. True when the procedure exists.
Private Function ModuleContainsProcedure(ByVal moduleText As String, _
                                         ByVal procedureName As String) As Boolean
    Dim subSignature As String
    Dim functionSignature As String

    If LenB(moduleText) = 0 Then Exit Function

    subSignature = "Sub " & procedureName & "("
    If InStr(1, moduleText, subSignature, vbTextCompare) > 0 Then
        ModuleContainsProcedure = True
        Exit Function
    End If

    functionSignature = "Function " & procedureName & "("
    ModuleContainsProcedure = (InStr(1, moduleText, functionSignature, vbTextCompare) > 0)
End Function

'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising pattern.

'@label:throw-error
'@sub-title Raise a ProjectError-coded runtime error
'@details
'Wrapper around Err.Raise that standardises the source to
'"CheckingOutput" for consistent stack traces.
'@param errNumb Long. ProjectError enumeration value.
'@param errorMessage String. Human-readable description.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "CheckingOutput", errorMessage
End Sub

'@section Interface
'===============================================================================
'@description Delegated members satisfying the ICheckingOutput contract.

Private Property Get ICheckingOutput_Wksh() As Worksheet
    Set ICheckingOutput_Wksh = Wksh()
End Property

Private Property Get ICheckingOutput_Name() As String
    ICheckingOutput_Name = Name()
End Property

'@label:print-output
'@sub-title Write checking values to a worksheet
'@details
'Coordinates the end-to-end rendering: validates the input array,
'initialises the worksheet, writes each checking bundle, appends the
'end marker, and records the final row position.
'@param checkTable BetterArray. Collection of IChecking instances to print.
'@throws ProjectError.ElementNotFound When checkTable is Nothing.
'@export
Private Sub ICheckingOutput_PrintOutput(ByVal checkTable As BetterArray)
    Dim currentRow As Long
    Dim writtenTitles As Collection
    Dim index As Long
    Dim previousEventState As Boolean
    Dim store As IHiddenNames

    If checkTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "PrintOutput requires a BetterArray of IChecking instances"
    End If

    previousEventState = Application.EnableEvents
    Application.EnableEvents = False
    ValidateCheckTable checkTable
    Set store = HiddenStore()

    On Error GoTo CleanExit

    currentRow = InitialiseWorksheet()
    Set writtenTitles = New Collection

    For index = checkTable.lowerBound To checkTable.upperBound
        WriteChecking checkTable.Item(index), currentRow, writtenTitles
    Next index

    '@Ignore AssignmentNotUsed
    currentRow = currentRow + 4
    WriteRow currentRow, formatTag:="level 4", textValue:="End of checkings " & Name(), parentTitle:=END_TAG
    store.SetValue ROW_MARKER_NAME, currentRow

CleanExit:
    Application.EnableEvents = previousEventState
    If Err.Number <> 0 Then ThrowError Err.Number, Err.description
End Sub

Private Sub ICheckingOutput_FilterWorksheet(Optional ByVal selectionValue As String = vbNullString, _
                                            Optional ByVal titleSelection As String = vbNullString)
    FilterWorksheet selectionValue, titleSelection
End Sub

Private Sub ICheckingOutput_EnsureWorksheetChangeHandler()
    EnsureWorksheetChangeHandler
End Sub
