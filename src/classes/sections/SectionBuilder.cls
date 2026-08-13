VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@PredeclaredId
'@Folder("Sections")
'@ModuleDescription("Consolidated section/subsection builder for HList and VList worksheets")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'Section layout modes handled by SectionBuilder. The enum lives in this class,
'so any class that depends on SectionBuilder resolves it. `test-registry.yml`
'records the same move under `folder: sections`.
Public Enum SectionBuilderMode
    SectionBuilderModeHList = 1
    SectionBuilderModeVList = 2
End Enum


'@class SectionBuilder
'@description
'Builds all the sections, subsections and variables of one linelist worksheet.
'Given a sheet name and a dictionary start row, Build scans the dictionary for
'the rows of that sheet, groups them by section and subsection, hands each
'variable to a VarWriter, and writes the merged section and subsection headers
'with their formatting. The layer byte picks the HList layout or the VList one.
'Entries filed along the way leave through HasCheckings and CheckingValues.
'
'ONE PASS OVER THE DICTIONARY
'-------------------------------------------------------------------------------
'`Build` reads the five dictionary columns it groups on into arrays, in five
'crossings, and every boundary scan after that runs in memory.
'
'ONE WRITER FOR THE SHEET
'-------------------------------------------------------------------------------
'`Build` creates the `VarWriter` and hands it to every subsection. The writer
'holds a hidden name store, the dictionary block, the design format and the
'translated messages, and all four cover a whole sheet.
'
'THE BOUNDARIES ARE LEFT BEHIND ON THE SHEET
'-------------------------------------------------------------------------------
'Every section `Build` lays out is also recorded in a `SectionMap`, so that a
'later session can act on a whole section without walking the dictionary again.
'The map is cleared at the head of the build and one block is added per section,
'holding the first and last position and the section title.
'
'The map is written on the main sheet alone. The printed companion carries the
'same columns and the CRF carries its own, and neither is offered a section
'action today.
'@depends LinelistSpecs, LLdictionary, LLFormat, DropdownLists, VarWriter,
'SectionMap, Checking
'@version 1.2 (2026-08-06)


Private Type TSectionBuilder
    layer As Byte
    specs As LinelistSpecs
    wksh As Worksheet
    printWksh As Worksheet
    crfWksh As Worksheet
    dropdown As DropdownLists
    customDropdown As DropdownLists
    cachedDict As LLdictionary
    cachedFormat As LLFormat
    secMap As SectionMap
    sectionNames As Variant
    subSectionNames As Variant
    sheetNames As Variant
    variableNames As Variant
    columnIndexes As Variant
    crfIndexes As Variant
    lastDataRow As Long
    hasChecks As Boolean
    sectionsPlaced As Long
    writer As VarWriter
    sealed As Boolean
End Type

Private this As TSectionBuilder
Private Const CLASS_NAME As String = "SectionBuilder"

'Checking support
Private internalChecks As Checking

' Row/column constants for section and subsection headers
Private Const HLIST_SEC_ROW As Long = 5
Private Const HLIST_SUBSEC_ROW As Long = 6
Private Const VLIST_SEC_COL As Long = 2
Private Const VLIST_SUBSEC_COL As Long = 3

' The CRF companion places a section four rows above the first variable of the
' section and a subsection three rows above.
Private Const CRF_SECTION_OFFSET As Long = 4
Private Const CRF_SUBSECTION_OFFSET As Long = 3
Private Const CRF_LAST_COLUMN As Long = 7


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new SectionBuilder instance
'@details
'Factory method on the predeclared instance. Validates that specs and
'wksh are not Nothing and that the layer names a SectionBuilderMode member,
'then initialises a new SectionBuilder with the given layer, specifications,
'and optional companion worksheets and dropdown managers.
'@param layer Byte. SectionBuilderMode enum value (HList=1, VList=2).
'@param specs LinelistSpecs. The linelist specifications object.
'@param wksh Worksheet. The target main worksheet.
'@param printWksh Optional Worksheet. Printed companion worksheet (HList only).
'@param crfWksh Optional Worksheet. CRF companion worksheet (HList only).
'@param dropdownObj Optional DropdownLists. Dropdown lists manager.
'@param customDropdownObj Optional DropdownLists. Custom dropdown lists manager.
'@return SectionBuilder. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When specs or wksh is Nothing.
'@throws ProjectError.InvalidArgument When layer names neither SectionBuilderMode member.
'@export
Public Function Create(ByVal layer As Byte, _
                       ByVal specs As LinelistSpecs, _
                       ByVal wksh As Worksheet, _
                       Optional ByVal printWksh As Worksheet = Nothing, _
                       Optional ByVal crfWksh As Worksheet = Nothing, _
                       Optional ByVal dropdownObj As DropdownLists = Nothing, _
                       Optional ByVal customDropdownObj As DropdownLists = Nothing) As SectionBuilder

    If specs Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Specifications object is required"
    If wksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Target worksheet is required"

    If (layer <> SectionBuilderModeHList) And (layer <> SectionBuilderModeVList) Then
        ThrowError ProjectError.InvalidArgument, _
                   "Layer should be SectionBuilderModeHList or SectionBuilderModeVList. Found: " & layer
    End If

    Dim newBuilder As SectionBuilder
    Set newBuilder = New SectionBuilder
    newBuilder.InternalLayer = layer
    Set newBuilder.InternalSpecs = specs
    Set newBuilder.InternalWksh = wksh
    If Not printWksh Is Nothing Then Set newBuilder.InternalPrintWksh = printWksh
    If Not crfWksh Is Nothing Then Set newBuilder.InternalCRFWksh = crfWksh
    If Not dropdownObj Is Nothing Then Set newBuilder.InternalDropdown = dropdownObj
    If Not customDropdownObj Is Nothing Then Set newBuilder.InternalCustomDropdown = customDropdownObj
    newBuilder.Seal
    Set Create = newBuilder
End Function

'@section Instance Sealing
'===============================================================================

'@label:seal
'@sub-title Seal the instance against further setup writes
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup-only setter against post-seal writes
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub


'@section Internal Setters
'===============================================================================
'@description Properties used by the factory only during instantiation.

'@label:internal-layer
'@prop-title Assign the layer mode
'@param layer Byte. SectionBuilderMode enum value.
Public Property Let InternalLayer(ByVal layer As Byte)
    GuardNotSealed "InternalLayer"
    this.layer = layer
End Property

'@label:internal-specs
'@prop-title Assign the linelist specifications
'@param specs LinelistSpecs. The specifications object.
Public Property Set InternalSpecs(ByVal specs As LinelistSpecs)
    GuardNotSealed "InternalSpecs"
    Set this.specs = specs
End Property

'@label:internal-wksh
'@prop-title Assign the target main worksheet
'@param wksh Worksheet. The main worksheet to build on.
Public Property Set InternalWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalWksh"
    Set this.wksh = wksh
End Property

'@label:internal-print-wksh
'@prop-title Assign the printed companion worksheet
'@param wksh Worksheet. The printed companion worksheet.
Public Property Set InternalPrintWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalPrintWksh"
    Set this.printWksh = wksh
End Property

'@label:internal-crf-wksh
'@prop-title Assign the CRF companion worksheet
'@param wksh Worksheet. The CRF companion worksheet.
Public Property Set InternalCRFWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalCRFWksh"
    Set this.crfWksh = wksh
End Property

'@label:internal-dropdown
'@prop-title Assign the dropdown lists manager
'@param dropdownObj DropdownLists. The dropdown manager.
Public Property Set InternalDropdown(ByVal dropdownObj As DropdownLists)
    GuardNotSealed "InternalDropdown"
    Set this.dropdown = dropdownObj
End Property

'@label:internal-custom-dropdown
'@prop-title Assign the custom dropdown lists manager
'@param dropdownObj DropdownLists. The custom dropdown manager.
Public Property Set InternalCustomDropdown(ByVal dropdownObj As DropdownLists)
    GuardNotSealed "InternalCustomDropdown"
    Set this.customDropdown = dropdownObj
End Property


'@section Cached Accessors
'===============================================================================
'@description Lazy-loaded accessors for frequently used dependency objects.

'@label:dictionary
'@prop-title Lazy-loaded dictionary accessor
'@return LLdictionary. The linelist dictionary.
Private Property Get Dictionary() As LLdictionary
    If this.cachedDict Is Nothing Then
        Set this.cachedDict = this.specs.Dictionary()
    End If
    Set Dictionary = this.cachedDict
End Property

'@label:design-format
'@prop-title Design format accessor
'@details
'Held rather than resolved per call: every section and every subsection asks
'for it.
'@return LLFormat. The formatting object from the specifications.
Private Property Get DesignFormat() As LLFormat
    If this.cachedFormat Is Nothing Then
        Set this.cachedFormat = this.specs.DesignFormat()
    End If
    Set DesignFormat = this.cachedFormat
End Property


'@section Entry Point
'===============================================================================

'@label:build
'@sub-title Build all sections, subsections, and variables for a sheet
'@details
'Reads the dictionary columns it groups on into memory, then scans them for
'rows matching sheetName starting from startRow. Groups consecutive rows by
'their "main section" value, delegates to BuildSubSections for subsection
'grouping and variable writing, and applies merged-cell section header
'formatting. Stops when the sheet name changes or the dictionary ends.
'
'Each section is also recorded in the sheet's SectionMap as it is laid out.
'The map is opened before the writer runs, because opening it reads every
'hidden name the sheet holds and the writer adds three to five of them per
'variable.
'@param sheetName String. The sheet name as it appears in the dictionary "sheet name" column.
'@param startRow Long. The first data row index in the dictionary for this sheet (1-based).
'@export
Public Sub Build(ByVal sheetName As String, ByVal startRow As Long)
    Dim currentRow As Long
    Dim sectionStartRow As Long
    Dim sectionEndRow As Long
    Dim currentSection As String
    Dim writer As VarWriter

    LoadDictionaryColumns
    OpenSectionMap

    Set writer = VarWriter.Create( _
        layer:=this.layer, _
        specs:=this.specs, _
        wksh:=this.wksh, _
        printWksh:=this.printWksh, _
        crfWksh:=this.crfWksh, _
        dropdownObj:=this.dropdown, _
        customDropdownObj:=this.customDropdown)

    'One writer serves the whole sheet, so the field is what lets the
    'caller read the variable count and the milestone record after Build.
    Set this.writer = writer

    currentRow = startRow

    ' Iterate through sections for this sheet
    Do While currentRow <= this.lastDataRow
        ' Check if still on the same sheet
        If TextAt(this.sheetNames, currentRow) <> sheetName Then Exit Do

        ' Start a new section
        currentSection = TextAt(this.sectionNames, currentRow)
        sectionStartRow = currentRow

        ' Find section end: scan until section name or sheet name changes
        sectionEndRow = FindSectionEnd(currentRow, currentSection, sheetName)

        ' Build all subsections within this section
        BuildSubSections sectionStartRow, sectionEndRow, sheetName, currentSection, writer

        ' Format the section header
        FormatSection currentSection, sectionStartRow, sectionEndRow

        ' Leave the boundaries of the section behind on the sheet
        RecordSection currentSection, sectionStartRow, sectionEndRow

        ' The section is laid out, so the milestone is true
        LogSectionPlaced currentSection, sheetName

        ' Move to next section
        currentRow = sectionEndRow + 1
    Loop

    HarvestWriterCheckings writer
End Sub


'@section Dictionary Reading
'===============================================================================

'@label:load-dictionary-columns
'@sub-title Read the columns the build groups on into memory
'@details
'Five crossings for the whole build. `crf index` is read under a guard because
'a dictionary that has not been prepared does not carry it, and the CRF
'companion is optional anyway.
Private Sub LoadDictionaryColumns()
    Dim dict As LLdictionary

    Set dict = Dictionary()
    this.lastDataRow = dict.Data.DataEndRow()

    this.sectionNames = ColumnValues(dict.DataRange("main section"))
    this.subSectionNames = ColumnValues(dict.DataRange("sub section"))
    this.sheetNames = ColumnValues(dict.DataRange("sheet name"))
    this.variableNames = ColumnValues(dict.DataRange("variable name"))
    this.columnIndexes = ColumnValues(dict.DataRange("column index"))

    this.crfIndexes = Empty
    If Not this.crfWksh Is Nothing Then
        'The column arrives with LLdictionary.Prepare and the lookup raises
        'without it. The suppression covers that one lookup.
        On Error Resume Next
            this.crfIndexes = ColumnValues(dict.DataRange("crf index"))
        On Error GoTo 0
    End If
End Sub

'@label:column-values
'@sub-title Read a single-column Range as a 1-based list of values
'@details
'`Range.Value` hands back a plain value for one cell and a two-dimensional
'array otherwise. This flattens both cases so callers see one shape.
'@param columnRng Range. A single-column block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function ColumnValues(ByVal columnRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbRows As Long

    If columnRng Is Nothing Then Exit Function

    nbRows = columnRng.Rows.Count
    ReDim flat(1 To nbRows)

    If nbRows = 1 Then
        flat(1) = columnRng.Cells(1, 1).Value
    Else
        rawValues = columnRng.Columns(1).Value
        For counter = 1 To nbRows
            flat(counter) = rawValues(counter, 1)
        Next counter
    End If

    ColumnValues = flat
End Function

'@label:text-at
'@sub-title Read one entry of a loaded column as text
'@details
'A dictionary cell can hold an error value or a Null, and `CStr` raises on
'both.
'@param values Variant. A 1-based list read by ColumnValues.
'@param rowNumber Long. The dictionary data row.
'@return String. The text, or an empty string.
Private Function TextAt(ByVal values As Variant, ByVal rowNumber As Long) As String
    Dim cellValue As Variant

    If IsEmpty(values) Then Exit Function
    If rowNumber < LBound(values) Or rowNumber > UBound(values) Then Exit Function

    cellValue = values(rowNumber)
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function

    TextAt = CStr(cellValue)
End Function

'@label:number-at
'@sub-title Read one entry of a loaded column as a whole number
'@details
'One rule for both index columns: a value that is empty or is not a number
'answers 0, and every caller tests for 0. `column index` and `crf index` used
'to be converted with a bare `CLng`, guarded at three of the eight sites that
'did it.
'@param values Variant. A 1-based list read by ColumnValues.
'@param rowNumber Long. The dictionary data row.
'@return Long. The stored number, or 0.
Private Function NumberAt(ByVal values As Variant, ByVal rowNumber As Long) As Long
    Dim storedText As String

    storedText = TextAt(values, rowNumber)
    If LenB(storedText) = 0 Then Exit Function
    If Not IsNumeric(storedText) Then Exit Function

    NumberAt = CLng(storedText)
End Function


'@section Section Boundary Detection
'===============================================================================

'@label:find-section-end
'@sub-title Find the last row of a section
'@details
'Starting from startRow, scans forward until the section name or sheet
'name changes, then returns the row just before the change. When the
'end of the data is reached without a change, returns the last data row.
'@param startRow Long. The row to start scanning from.
'@param sectionName String. The current section name to match.
'@param sheetName String. The current sheet name to match.
'@return Long. The 1-based row index of the last row in the section.
Private Function FindSectionEnd(ByVal startRow As Long, _
                                 ByVal sectionName As String, _
                                 ByVal sheetName As String) As Long
    Dim rowNumber As Long
    rowNumber = startRow

    Do While rowNumber < this.lastDataRow
        rowNumber = rowNumber + 1
        If (TextAt(this.sectionNames, rowNumber) <> sectionName) Or _
           (TextAt(this.sheetNames, rowNumber) <> sheetName) Then
            FindSectionEnd = rowNumber - 1
            Exit Function
        End If
    Loop

    FindSectionEnd = rowNumber
End Function


'@section Section Map
'===============================================================================
'@description
'The boundaries of every section, left on the sheet for a later session. A
'section is a run of consecutive dictionary rows, and the map holds the first
'and last position of each run with the title alongside them. The reader of the
'map starts from a selected cell, so the position range is what it looks a
'section up by.

'@label:open-section-map
'@sub-title Open the sheet's section map and empty it
'@details
'Called at the head of the build, before the writer adds its own names, and the
'clear is what stops a sheet rebuilt with fewer sections than last time from
'keeping the tail of the old map.
'@depends SectionMap
Private Sub OpenSectionMap()
    Set this.secMap = SectionMap.Create(this.wksh)
    this.secMap.Clear
End Sub

'@label:record-section
'@sub-title Record the boundaries of one section in the map
'@details
'The positions are the same `column index` values the headers were laid out on:
'a column number on an HList sheet, a row number on a VList one. A section
'whose first or last row carries no position is skipped, for the same reason
'FormatSection leaves it alone.
'@param sectionName String. The section title.
'@param startRow Long. The first dictionary row of the section.
'@param endRow Long. The last dictionary row of the section.
Private Sub RecordSection(ByVal sectionName As String, _
                           ByVal startRow As Long, _
                           ByVal endRow As Long)
    Dim startPos As Long
    Dim endPos As Long

    If this.secMap Is Nothing Then Exit Sub

    startPos = NumberAt(this.columnIndexes, startRow)
    endPos = NumberAt(this.columnIndexes, endRow)
    If (startPos <= 0) Or (endPos < startPos) Then Exit Sub

    this.secMap.Add sectionName, startPos, endPos
End Sub


'@section Subsection Processing
'===============================================================================

'@label:build-sub-sections
'@sub-title Build all subsections within a section range
'@details
'Iterates through dictionary rows from sectionStart to sectionEnd,
'grouping consecutive rows by the "sub section" value. For each subsection
'group, delegates variable writing to BuildVariables and then applies
'subsection header formatting via FormatSubSection.
'@param sectionStart Long. The first row of the section.
'@param sectionEnd Long. The last row of the section.
'@param sheetName String. The current sheet name.
'@param sectionName String. The enclosing section name.
'@param writer VarWriter. The writer that serves the whole sheet.
Private Sub BuildSubSections(ByVal sectionStart As Long, _
                              ByVal sectionEnd As Long, _
                              ByVal sheetName As String, _
                              ByVal sectionName As String, _
                              ByVal writer As VarWriter)
    Dim currentRow As Long
    Dim subSecStart As Long
    Dim subSecEnd As Long
    Dim currentSubSection As String

    currentRow = sectionStart

    Do While currentRow <= sectionEnd
        currentSubSection = TextAt(this.subSectionNames, currentRow)
        subSecStart = currentRow

        ' Find subsection end
        subSecEnd = FindSubSectionEnd(currentRow, currentSubSection, sectionName, _
                                       sheetName, sectionEnd)

        ' Build variables within this subsection
        BuildVariables subSecStart, subSecEnd, writer

        ' Format subsection header
        FormatSubSection currentSubSection, subSecStart, subSecEnd

        currentRow = subSecEnd + 1
    Loop
End Sub

'@label:find-sub-section-end
'@sub-title Find the last row of a subsection
'@details
'Starting from startRow, scans forward until the subsection name,
'section name, or sheet name changes, then returns the row just before
'the change. Constrained by maxRow (the section boundary).
'@param startRow Long. The row to start scanning from.
'@param subSectionName String. The current subsection name to match.
'@param sectionName String. The current section name to match.
'@param sheetName String. The current sheet name to match.
'@param maxRow Long. The last row of the enclosing section.
'@return Long. The 1-based row index of the last row in the subsection.
Private Function FindSubSectionEnd(ByVal startRow As Long, _
                                    ByVal subSectionName As String, _
                                    ByVal sectionName As String, _
                                    ByVal sheetName As String, _
                                    ByVal maxRow As Long) As Long
    Dim rowNumber As Long
    rowNumber = startRow

    Do While rowNumber < maxRow
        rowNumber = rowNumber + 1
        If (TextAt(this.subSectionNames, rowNumber) <> subSectionName) Or _
           (TextAt(this.sectionNames, rowNumber) <> sectionName) Or _
           (TextAt(this.sheetNames, rowNumber) <> sheetName) Then
            FindSubSectionEnd = rowNumber - 1
            Exit Function
        End If
    Loop

    FindSubSectionEnd = rowNumber
End Function


'@section Variable Building
'===============================================================================

'@label:build-variables
'@sub-title Build all variables within a subsection range
'@details
'Walks the subsection and hands each variable name to the writer `Build`
'created for the sheet.
'@param subSecStart Long. The first dictionary row of the subsection.
'@param subSecEnd Long. The last dictionary row of the subsection.
'@param writer VarWriter. The writer that serves the whole sheet.
'@depends VarWriter
Private Sub BuildVariables(ByVal subSecStart As Long, _
                            ByVal subSecEnd As Long, _
                            ByVal writer As VarWriter)
    Dim counter As Long
    Dim varName As String

    For counter = subSecStart To subSecEnd
        varName = TextAt(this.variableNames, counter)
        writer.WriteVariable varName
    Next counter
End Sub


'@section Section Formatting
'===============================================================================

'@label:format-section
'@sub-title Format a section header
'@details
'Dispatches to FormatHListSection or FormatVListSection based on the
'current layer.
'@param sectionName String. The section title to display.
'@param startRow Long. The first dictionary row of the section.
'@param endRow Long. The last dictionary row of the section.
Private Sub FormatSection(ByVal sectionName As String, _
                           ByVal startRow As Long, _
                           ByVal endRow As Long)
    If this.layer = SectionBuilderModeHList Then
        FormatHListSection sectionName, startRow, endRow
    Else
        FormatVListSection sectionName, startRow, endRow
    End If
End Sub

'@label:format-hlist-section
'@sub-title Format an HList section header
'@details
'Merges cells horizontally across the column range at HLIST_SEC_ROW on
'the main sheet, writes the section name, and applies HListSection
'formatting. Repeats for the printed companion when available. Also
'writes a CRF section header when a CRF worksheet is present, using the
'first variable's "crf index" to calculate the CRF row offset. The CRF line
'is read inside the CRF guard, so a linelist with no CRF companion pays
'nothing for it.
'@param sectionName String. The section title to display.
'@param startRow Long. The first dictionary row of the section.
'@param endRow Long. The last dictionary row of the section.
Private Sub FormatHListSection(ByVal sectionName As String, _
                                ByVal startRow As Long, _
                                ByVal endRow As Long)
    Dim design As LLFormat
    Dim startCol As Long
    Dim endCol As Long
    Dim crfLine As Long
    Dim rng As Range

    Set design = DesignFormat()

    startCol = NumberAt(this.columnIndexes, startRow)
    endCol = NumberAt(this.columnIndexes, endRow)
    If (startCol <= 0) Or (endCol <= 0) Then Exit Sub

    ' Main sheet section
    With this.wksh
        Set rng = .Range(.Cells(HLIST_SEC_ROW, startCol), .Cells(HLIST_SEC_ROW, endCol))
    End With
    rng.Cells(1, 1).Value = sectionName
    design.ApplyFormat rng, scope:=HListSection

    ' Printed sheet section
    If Not this.printWksh Is Nothing Then
        With this.printWksh
            Set rng = .Range(.Cells(HLIST_SEC_ROW, startCol), .Cells(HLIST_SEC_ROW, endCol))
        End With
        rng.Cells(1, 1).Value = sectionName
        design.ApplyFormat rng, scope:=HListSection
    End If

    ' CRF sheet section
    If (sectionName <> vbNullString) And (Not this.crfWksh Is Nothing) Then
        crfLine = NumberAt(this.crfIndexes, startRow)
        If crfLine <= CRF_SECTION_OFFSET Then Exit Sub

        With this.crfWksh
            Set rng = .Range(.Cells(crfLine - CRF_SECTION_OFFSET, 1), _
                             .Cells(crfLine - CRF_SECTION_OFFSET, CRF_LAST_COLUMN))
        End With
        rng.Cells(1, 1).Value = sectionName
        design.ApplyFormat rng, scope:=HListCRFSection
    End If
End Sub

'@label:format-vlist-section
'@sub-title Format a VList section header
'@details
'Merges cells vertically across the row range at VLIST_SEC_COL on the
'main sheet, writes the section name, and applies VListSection
'formatting. No companion sheets are involved for VList layers.
'@param sectionName String. The section title to display.
'@param startRow Long. The first dictionary row of the section.
'@param endRow Long. The last dictionary row of the section.
Private Sub FormatVListSection(ByVal sectionName As String, _
                                ByVal startRow As Long, _
                                ByVal endRow As Long)
    Dim design As LLFormat
    Dim startIdx As Long
    Dim endIdx As Long
    Dim rng As Range

    Set design = DesignFormat()

    startIdx = NumberAt(this.columnIndexes, startRow)
    endIdx = NumberAt(this.columnIndexes, endRow)
    If (startIdx <= 0) Or (endIdx <= 0) Then Exit Sub

    With this.wksh
        Set rng = .Range(.Cells(startIdx, VLIST_SEC_COL), .Cells(endIdx, VLIST_SEC_COL))
    End With

    rng.Cells(1, 1).Value = sectionName
    design.ApplyFormat rng, scope:=VListSection
End Sub


'@section Subsection Formatting
'===============================================================================

'@label:format-sub-section
'@sub-title Format a subsection header
'@details
'Dispatches to FormatHListSubSection or FormatVListSubSection based on
'the current layer.
'@param subSectionName String. The subsection title to display.
'@param startRow Long. The first dictionary row of the subsection.
'@param endRow Long. The last dictionary row of the subsection.
Private Sub FormatSubSection(ByVal subSectionName As String, _
                              ByVal startRow As Long, _
                              ByVal endRow As Long)
    If this.layer = SectionBuilderModeHList Then
        FormatHListSubSection subSectionName, startRow, endRow
    Else
        FormatVListSubSection subSectionName, startRow, endRow
    End If
End Sub

'@label:format-hlist-sub-section
'@sub-title Format an HList subsection header
'@details
'Merges cells horizontally across the column range at HLIST_SUBSEC_ROW
'on the main sheet, writes the subsection name, and applies
'HListSubSection formatting. Repeats for the printed companion when
'available. Also writes a CRF subsection header when a CRF worksheet
'is present.
'@param subSectionName String. The subsection title to display.
'@param startRow Long. The first dictionary row of the subsection.
'@param endRow Long. The last dictionary row of the subsection.
Private Sub FormatHListSubSection(ByVal subSectionName As String, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long)
    Dim design As LLFormat
    Dim startCol As Long
    Dim endCol As Long
    Dim crfLine As Long
    Dim rng As Range

    Set design = DesignFormat()

    startCol = NumberAt(this.columnIndexes, startRow)
    endCol = NumberAt(this.columnIndexes, endRow)
    If (startCol <= 0) Or (endCol <= 0) Then Exit Sub

    ' Main sheet subsection
    With this.wksh
        Set rng = .Range(.Cells(HLIST_SUBSEC_ROW, startCol), .Cells(HLIST_SUBSEC_ROW, endCol))
    End With
    rng.Cells(1, 1).Value = subSectionName
    design.ApplyFormat rng, scope:=HListSubSection

    ' Printed sheet subsection
    If Not this.printWksh Is Nothing Then
        With this.printWksh
            Set rng = .Range(.Cells(HLIST_SUBSEC_ROW, startCol), .Cells(HLIST_SUBSEC_ROW, endCol))
        End With
        rng.Cells(1, 1).Value = subSectionName
        design.ApplyFormat rng, scope:=HListSubSection
    End If

    ' CRF sheet subsection
    If (subSectionName <> vbNullString) And (Not this.crfWksh Is Nothing) Then
        crfLine = NumberAt(this.crfIndexes, startRow)
        If crfLine <= CRF_SUBSECTION_OFFSET Then Exit Sub

        With this.crfWksh
            Set rng = .Range(.Cells(crfLine - CRF_SUBSECTION_OFFSET, 1), _
                             .Cells(crfLine - CRF_SUBSECTION_OFFSET, CRF_LAST_COLUMN))
        End With
        rng.Cells(1, 1).Value = subSectionName
        design.ApplyFormat rng, scope:=HListCRFSubSection
    End If
End Sub

'@label:format-vlist-sub-section
'@sub-title Format a VList subsection header
'@details
'Merges cells vertically across the row range at VLIST_SUBSEC_COL on
'the main sheet, writes the subsection name, and applies
'VListSubSection formatting.
'@param subSectionName String. The subsection title to display.
'@param startRow Long. The first dictionary row of the subsection.
'@param endRow Long. The last dictionary row of the subsection.
Private Sub FormatVListSubSection(ByVal subSectionName As String, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long)
    Dim design As LLFormat
    Dim startIdx As Long
    Dim endIdx As Long
    Dim rng As Range

    Set design = DesignFormat()

    startIdx = NumberAt(this.columnIndexes, startRow)
    endIdx = NumberAt(this.columnIndexes, endRow)
    If (startIdx <= 0) Or (endIdx <= 0) Then Exit Sub

    With this.wksh
        Set rng = .Range(.Cells(startIdx, VLIST_SUBSEC_COL), .Cells(endIdx, VLIST_SUBSEC_COL))
    End With

    rng.Cells(1, 1).Value = subSectionName
    design.ApplyFormat rng, scope:=VListSubSection
End Sub


'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise an error with the class name as source
'@details
'Raises the number it is given with the class name as the source. Every guard
'in this class names a `ProjectError` member.
'@param errNumber Long. The ProjectError value to raise.
'@param errorMessage String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, "Error : " & errorMessage
End Sub


'@section Checkings
'===============================================================================
'@description
'The writer files one entry per problem, keyed by worksheet and variable, and
'`Build` merges them once when the sheet is done. The merge used to run per
'subsection against a writer whose keys started at 0 in every instance, so the
'second subsection of a sheet re-added key "0" and `Checking.Add` stopped the
'build.

'@label:harvest-writer-checkings
'@sub-title Take the writer's entries once the sheet is built
'@param writer VarWriter. The writer that served the sheet.
Private Sub HarvestWriterCheckings(ByVal writer As VarWriter)
    If writer Is Nothing Then Exit Sub
    If Not writer.HasCheckings Then Exit Sub

    If internalChecks Is Nothing Then
        Set internalChecks = writer.CheckingValues.Clone()
    Else
        internalChecks.Append writer.CheckingValues
    End If

    this.hasChecks = True
End Sub

'@label:log-section-placed
'@sub-title File the milestone of one section laid out
'@details
'A sheet holds a few sections, so these entries go to the worksheet with
'the problems. The per-variable entries are the ones held back, and they
'travel through `MilestoneValues`.
'@param sectionName String. The section title.
'@param sheetName String. The dictionary sheet the section belongs to.
Private Sub LogSectionPlaced(ByVal sectionName As String, ByVal sheetName As String)
    Dim keyRoot As String
    Dim candidate As String
    Dim suffix As Long
    Dim title As String

    If internalChecks Is Nothing Then Set internalChecks = Checking.Create("Sections")

    'A run of dictionary rows with no section title is still a section the
    'build laid out, and the map records it. Reading "section  placed" makes
    'the report look broken, so the empty title is named.
    title = sectionName
    If LenB(Trim$(title)) = 0 Then title = "(untitled)"

    keyRoot = sheetName & "!" & title
    candidate = keyRoot
    Do While internalChecks.KeyExists(candidate)
        suffix = suffix + 1
        candidate = keyRoot & " #" & CStr(suffix)
    Loop

    internalChecks.Add candidate, _
                       "section " & title & " placed on " & sheetName, _
                       checkingInfo
    this.sectionsPlaced = this.sectionsPlaced + 1
    this.hasChecks = True
End Sub

Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@label:has-milestones
'@prop-title Whether the writer filed any per-variable milestone
'@return Boolean. True once a variable of this sheet was written.
'@export
Public Property Get HasMilestones() As Boolean
    If this.writer Is Nothing Then Exit Property
    HasMilestones = this.writer.HasMilestones
End Property

'@label:milestone-values
'@prop-title The writer's per-variable record
'@details
'Handed on as the writer holds it. One writer serves the whole sheet, so
'no merge is owed here, and the store stays out of the __check worksheet.
'@return Checking. The milestone entries, or Nothing.
'@export
Public Property Get MilestoneValues() As Checking
    If this.writer Is Nothing Then Exit Property
    Set MilestoneValues = this.writer.MilestoneValues
End Property

'@label:variables-written
'@prop-title How many variables the writer put on the sheet
'@return Long. The number of variables written.
'@export
Public Property Get VariablesWritten() As Long
    If this.writer Is Nothing Then Exit Property
    VariablesWritten = this.writer.VariablesWritten
End Property

'@label:sections-placed
'@prop-title How many sections were laid out on the sheet
'@return Long. The number of sections placed.
'@export
Public Property Get SectionsPlaced() As Long
    SectionsPlaced = this.sectionsPlaced
End Property
