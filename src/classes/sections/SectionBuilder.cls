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
Implements ISectionBuilder

'@class SectionBuilder
'@description
'Builds all sections, subsections, and variables for a single linelist
'worksheet. Given a sheet name and dictionary start row, the builder
'scans the dictionary for matching rows, groups them by section and
'subsection, delegates variable writing to VarWriter, and applies
'merged-cell section and subsection header formatting. Supports both
'HList (horizontal) and VList (vertical) layouts via the layer byte.
'@depends ILinelistSpecs, ILLdictionary, ILLVariables, ILLFormat,
'IDropdownLists, VarWriter, LLVariables
'@version 1.0 (2026-02-10)



Private Type TSectionBuilder
    layer As Byte
    specs As ILinelistSpecs
    wksh As Worksheet
    printWksh As Worksheet
    crfWksh As Worksheet
    dropdown As IDropdownLists
    customDropdown As IDropdownLists
    cachedDict As ILLdictionary
    cachedVars As ILLVariables
End Type

Private this As TSectionBuilder
Private Const CLASS_NAME As String = "SectionBuilder"

' Row/column constants for section and subsection headers
Private Const HLIST_SEC_ROW As Long = 5
Private Const HLIST_SUBSEC_ROW As Long = 6
Private Const VLIST_SEC_COL As Long = 2
Private Const VLIST_SUBSEC_COL As Long = 3
Private Const CRF_SHEET_PREFIX As String = "crf_"


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new SectionBuilder instance
'@details
'Factory method on the predeclared instance. Validates that specs and
'wksh are not Nothing, then initialises a new SectionBuilder with the
'given layer, specifications, and optional companion worksheets and
'dropdown managers. Returns the instance through the ISectionBuilder
'interface.
'@param layer Byte. SectionBuilderMode enum value (HList=1, VList=2).
'@param specs ILinelistSpecs. The linelist specifications object.
'@param wksh Worksheet. The target main worksheet.
'@param printWksh Optional Worksheet. Printed companion worksheet (HList only).
'@param crfWksh Optional Worksheet. CRF companion worksheet (HList only).
'@param dropdownObj Optional IDropdownLists. Dropdown lists manager.
'@param customDropdownObj Optional IDropdownLists. Custom dropdown lists manager.
'@return ISectionBuilder. A fully initialised instance ready for use.
'@throws VbObjectError+516 When specs or wksh is Nothing.
'@export
Public Function Create(ByVal layer As Byte, _
                       ByVal specs As ILinelistSpecs, _
                       ByVal wksh As Worksheet, _
                       Optional ByVal printWksh As Worksheet = Nothing, _
                       Optional ByVal crfWksh As Worksheet = Nothing, _
                       Optional ByVal dropdownObj As IDropdownLists = Nothing, _
                       Optional ByVal customDropdownObj As IDropdownLists = Nothing) As ISectionBuilder

    If specs Is Nothing Then ThrowError "Specifications object is required"
    If wksh Is Nothing Then ThrowError "Target worksheet is required"

    With New SectionBuilder
        .InternalLayer = layer
        Set .InternalSpecs = specs
        Set .InternalWksh = wksh
        If Not printWksh Is Nothing Then Set .InternalPrintWksh = printWksh
        If Not crfWksh Is Nothing Then Set .InternalCRFWksh = crfWksh
        If Not dropdownObj Is Nothing Then Set .InternalDropdown = dropdownObj
        If Not customDropdownObj Is Nothing Then Set .InternalCustomDropdown = customDropdownObj
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to the ISectionBuilder interface,
'used by the Create factory to return the newly built object.
'@return ISectionBuilder. The current instance.
Public Property Get Self() As ISectionBuilder
    Set Self = Me
End Property


'@section Internal Setters
'===============================================================================
'@description Properties used by the factory only during instantiation.

'@label:internal-layer
'@prop-title Assign the layer mode
'@param layer Byte. SectionBuilderMode enum value.
Public Property Let InternalLayer(ByVal layer As Byte)
    this.layer = layer
End Property

'@label:internal-specs
'@prop-title Assign the linelist specifications
'@param specs ILinelistSpecs. The specifications object.
Public Property Set InternalSpecs(ByVal specs As ILinelistSpecs)
    Set this.specs = specs
End Property

'@label:internal-wksh
'@prop-title Assign the target main worksheet
'@param wksh Worksheet. The main worksheet to build on.
Public Property Set InternalWksh(ByVal wksh As Worksheet)
    Set this.wksh = wksh
End Property

'@label:internal-print-wksh
'@prop-title Assign the printed companion worksheet
'@param wksh Worksheet. The printed companion worksheet.
Public Property Set InternalPrintWksh(ByVal wksh As Worksheet)
    Set this.printWksh = wksh
End Property

'@label:internal-crf-wksh
'@prop-title Assign the CRF companion worksheet
'@param wksh Worksheet. The CRF companion worksheet.
Public Property Set InternalCRFWksh(ByVal wksh As Worksheet)
    Set this.crfWksh = wksh
End Property

'@label:internal-dropdown
'@prop-title Assign the dropdown lists manager
'@param dropdownObj IDropdownLists. The dropdown manager.
Public Property Set InternalDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.dropdown = dropdownObj
End Property

'@label:internal-custom-dropdown
'@prop-title Assign the custom dropdown lists manager
'@param dropdownObj IDropdownLists. The custom dropdown manager.
Public Property Set InternalCustomDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.customDropdown = dropdownObj
End Property


'@section Cached Accessors
'===============================================================================
'@description Lazy-loaded accessors for frequently used dependency objects.

'@label:dictionary
'@prop-title Lazy-loaded dictionary accessor
'@details
'Returns the cached ILLdictionary instance. On first access, retrieves
'it from the specifications object and caches it for subsequent calls.
'@return ILLdictionary. The linelist dictionary.
Private Property Get Dictionary() As ILLdictionary
    If this.cachedDict Is Nothing Then
        Set this.cachedDict = this.specs.Dictionary()
    End If
    Set Dictionary = this.cachedDict
End Property

'@label:variables
'@prop-title Lazy-loaded variables accessor
'@details
'Returns the cached ILLVariables instance. On first access, creates it
'from the dictionary via LLVariables.Create and caches it.
'@return ILLVariables. The variables helper for dictionary lookups.
Private Property Get Variables() As ILLVariables
    If this.cachedVars Is Nothing Then
        Set this.cachedVars = LLVariables.Create(Dictionary())
    End If
    Set Variables = this.cachedVars
End Property

'@label:design-format
'@prop-title Design format accessor
'@return ILLFormat. The formatting object from the specifications.
Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = this.specs.DesignFormat()
End Property


'@section Entry Point
'===============================================================================

'@label:build
'@sub-title Build all sections, subsections, and variables for a sheet
'@details
'Scans the dictionary for rows matching sheetName starting from
'startRow. Groups consecutive rows by their "main section" column value,
'then delegates to BuildSubSections for subsection grouping and variable
'writing. After all variables in a section are written, applies
'merged-cell section header formatting via FormatSection. Stops when the
'sheet name changes or the dictionary ends.
'@param sheetName String. The sheet name as it appears in the dictionary "sheet name" column.
'@param startRow Long. The first data row index in the dictionary for this sheet (1-based).
'@export
Public Sub Build(ByVal sheetName As String, ByVal startRow As Long)
    Dim dict As ILLdictionary
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim dataEndRow As Long
    Dim currentRow As Long
    Dim sectionStartRow As Long
    Dim sectionEndRow As Long
    Dim currentSection As String

    Set dict = Dictionary()
    Set sectionRng = dict.DataRange("main section")
    Set sheetRng = dict.DataRange("sheet name")
    dataEndRow = dict.Data.DataEndRow()

    currentRow = startRow

    ' Iterate through sections for this sheet
    Do While currentRow <= dataEndRow
        ' Check if still on the same sheet
        If sheetRng.Cells(currentRow, 1).Value <> sheetName Then Exit Do

        ' Start a new section
        currentSection = sectionRng.Cells(currentRow, 1).Value
        sectionStartRow = currentRow

        ' Find section end: scan until section name or sheet name changes
        sectionEndRow = FindSectionEnd(currentRow, currentSection, sheetName, _
                                        sectionRng, sheetRng, dataEndRow)

        ' Build all subsections within this section
        BuildSubSections sectionStartRow, sectionEndRow, sheetName

        ' Format the section header
        FormatSection currentSection, sectionStartRow, sectionEndRow, sheetName

        ' Move to next section
        currentRow = sectionEndRow + 1
    Loop
End Sub


'@section Section Boundary Detection
'===============================================================================

'@label:find-section-end
'@sub-title Find the last row of a section
'@details
'Starting from startRow, scans forward until the section name or sheet
'name changes, then returns the row just before the change. When the
'end of the data range is reached without a change, returns the last
'data row.
'@param startRow Long. The row to start scanning from.
'@param sectionName String. The current section name to match.
'@param sheetName String. The current sheet name to match.
'@param sectionRng Range. The "main section" data column range.
'@param sheetRng Range. The "sheet name" data column range.
'@param dataEndRow Long. The last row of data in the dictionary.
'@return Long. The 1-based row index of the last row in the section.
Private Function FindSectionEnd(ByVal startRow As Long, _
                                 ByVal sectionName As String, _
                                 ByVal sheetName As String, _
                                 ByVal sectionRng As Range, _
                                 ByVal sheetRng As Range, _
                                 ByVal dataEndRow As Long) As Long
    Dim rowNumber As Long
    rowNumber = startRow

    Do While rowNumber < dataEndRow
        rowNumber = rowNumber + 1
        If (sectionRng.Cells(rowNumber, 1).Value <> sectionName) Or _
           (sheetRng.Cells(rowNumber, 1).Value <> sheetName) Then
            FindSectionEnd = rowNumber - 1
            Exit Function
        End If
    Loop

    FindSectionEnd = rowNumber
End Function


'@section Subsection Processing
'===============================================================================

'@label:build-sub-sections
'@sub-title Build all subsections within a section range
'@details
'Iterates through dictionary rows from sectionStart to sectionEnd,
'grouping consecutive rows by the "sub section" column value. For each
'subsection group, delegates variable writing to BuildVariables and then
'applies subsection header formatting via FormatSubSection.
'@param sectionStart Long. The first row of the section.
'@param sectionEnd Long. The last row of the section.
'@param sheetName String. The current sheet name.
Private Sub BuildSubSections(ByVal sectionStart As Long, _
                              ByVal sectionEnd As Long, _
                              ByVal sheetName As String)
    Dim dict As ILLdictionary
    Dim subSectionRng As Range
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim currentRow As Long
    Dim subSecStart As Long
    Dim subSecEnd As Long
    Dim currentSubSection As String

    Set dict = Dictionary()
    Set subSectionRng = dict.DataRange("sub section")
    Set sectionRng = dict.DataRange("main section")
    Set sheetRng = dict.DataRange("sheet name")

    currentRow = sectionStart

    Do While currentRow <= sectionEnd
        currentSubSection = subSectionRng.Cells(currentRow, 1).Value
        subSecStart = currentRow

        ' Find subsection end
        subSecEnd = FindSubSectionEnd(currentRow, currentSubSection, _
                                       sectionRng.Cells(sectionStart, 1).Value, _
                                       sheetName, subSectionRng, sectionRng, _
                                       sheetRng, sectionEnd)

        ' Build variables within this subsection
        BuildVariables subSecStart, subSecEnd

        ' Format subsection header
        FormatSubSection currentSubSection, subSecStart, subSecEnd, sheetName

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
'@param subSectionRng Range. The "sub section" data column range.
'@param sectionRng Range. The "main section" data column range.
'@param sheetRng Range. The "sheet name" data column range.
'@param maxRow Long. The last row of the enclosing section.
'@return Long. The 1-based row index of the last row in the subsection.
Private Function FindSubSectionEnd(ByVal startRow As Long, _
                                    ByVal subSectionName As String, _
                                    ByVal sectionName As String, _
                                    ByVal sheetName As String, _
                                    ByVal subSectionRng As Range, _
                                    ByVal sectionRng As Range, _
                                    ByVal sheetRng As Range, _
                                    ByVal maxRow As Long) As Long
    Dim rowNumber As Long
    rowNumber = startRow

    Do While rowNumber < maxRow
        rowNumber = rowNumber + 1
        If (subSectionRng.Cells(rowNumber, 1).Value <> subSectionName) Or _
           (sectionRng.Cells(rowNumber, 1).Value <> sectionName) Or _
           (sheetRng.Cells(rowNumber, 1).Value <> sheetName) Then
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
'Creates a single VarWriter instance configured with the current layer,
'specifications, worksheets, and dropdown managers, then iterates from
'subSecStart to subSecEnd, calling WriteVariable for each variable name
'found in the dictionary. Reusing one VarWriter across the subsection
'avoids repeated factory overhead.
'@param subSecStart Long. The first dictionary row of the subsection.
'@param subSecEnd Long. The last dictionary row of the subsection.
'@depends VarWriter
Private Sub BuildVariables(ByVal subSecStart As Long, ByVal subSecEnd As Long)
    Dim dict As ILLdictionary
    Dim varRng As Range
    Dim counter As Long
    Dim varName As String
    Dim writer As IVarWriter

    Set dict = Dictionary()
    Set varRng = dict.DataRange("variable name")

    ' Create one VarWriter for this batch (reused for all variables)
    Set writer = VarWriter.Create( _
        layer:=this.layer, _
        specs:=this.specs, _
        wksh:=this.wksh, _
        printWksh:=this.printWksh, _
        crfWksh:=this.crfWksh, _
        dropdownObj:=this.dropdown, _
        customDropdownObj:=this.customDropdown)

    For counter = subSecStart To subSecEnd
        varName = varRng.Cells(counter, 1).Value
        writer.WriteVariable varName
    Next
End Sub


'@section Section Formatting
'===============================================================================

'@label:format-section
'@sub-title Format a section header
'@details
'Dispatches to FormatHListSection or FormatVListSection based on the
'current layer. Resolves the first and last variable names for the
'section range to calculate column or row indices for the merged header.
'@param sectionName String. The section title to display.
'@param startRow Long. The first dictionary row of the section.
'@param endRow Long. The last dictionary row of the section.
'@param sheetName String. The current sheet name.
Private Sub FormatSection(ByVal sectionName As String, _
                           ByVal startRow As Long, _
                           ByVal endRow As Long, _
                           ByVal sheetName As String)
    Dim vars As ILLVariables
    Dim varRng As Range
    Dim dict As ILLdictionary
    Dim design As ILLFormat
    Dim firstVarName As String
    Dim lastVarName As String
    Dim rng As Range

    Set dict = Dictionary()
    Set vars = Variables()
    Set design = DesignFormat()
    Set varRng = dict.DataRange("variable name")

    firstVarName = varRng.Cells(startRow, 1).Value
    lastVarName = varRng.Cells(endRow, 1).Value

    If this.layer = SectionBuilderModeHList Then
        FormatHListSection sectionName, firstVarName, lastVarName, sheetName
    Else
        FormatVListSection sectionName, firstVarName, lastVarName, sheetName
    End If
End Sub

'@label:format-hlist-section
'@sub-title Format an HList section header
'@details
'Merges cells horizontally across the column range at HLIST_SEC_ROW on
'the main sheet, writes the section name, and applies HListSection
'formatting. Repeats for the printed companion when available. Also
'writes a CRF section header when a CRF worksheet is present, using the
'first variable's "crf index" to calculate the CRF row offset.
'@param sectionName String. The section title to display.
'@param firstVarName String. The first variable name in the section.
'@param lastVarName String. The last variable name in the section.
'@param sheetName String. The current sheet name.
Private Sub FormatHListSection(ByVal sectionName As String, _
                                ByVal firstVarName As String, _
                                ByVal lastVarName As String, _
                                ByVal sheetName As String)
    Dim vars As ILLVariables
    Dim design As ILLFormat
    Dim startCol As Long
    Dim endCol As Long
    Dim crfLine As Long
    Dim rng As Range
    Dim sh As Worksheet

    Set vars = Variables()
    Set design = DesignFormat()

    startCol = CLng(vars.Value(varName:=firstVarName, colName:="column index"))
    endCol = CLng(vars.Value(varName:=lastVarName, colName:="column index"))
    crfLine = CLng(vars.Value(colName:="crf index", varName:=firstVarName))

    ' Main sheet section
    Set sh = this.wksh
    With sh
        Set rng = .Range(.Cells(HLIST_SEC_ROW, startCol), .Cells(HLIST_SEC_ROW, endCol))
        rng.Cells(1, 1).Value = sectionName
    End With
    design.ApplyFormat rng, scope:=HListSection

    ' Printed sheet section
    If Not this.printWksh Is Nothing Then
        Set sh = this.printWksh
        With sh
            Set rng = .Range(.Cells(HLIST_SEC_ROW, startCol), .Cells(HLIST_SEC_ROW, endCol))
            rng.Cells(1, 1).Value = sectionName
        End With
        design.ApplyFormat rng, scope:=HListSection
    End If

    ' CRF sheet section
    If (sectionName <> vbNullString) And (Not this.crfWksh Is Nothing) Then
        Set sh = this.crfWksh
        With sh
            Set rng = .Range(.Cells(crfLine - 4, 1), .Cells(crfLine - 4, 7))
            rng.Cells(1, 1).Value = sectionName
        End With
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
'@param firstVarName String. The first variable name in the section.
'@param lastVarName String. The last variable name in the section.
'@param sheetName String. The current sheet name.
Private Sub FormatVListSection(ByVal sectionName As String, _
                                ByVal firstVarName As String, _
                                ByVal lastVarName As String, _
                                ByVal sheetName As String)
    Dim vars As ILLVariables
    Dim design As ILLFormat
    Dim startIdx As Long
    Dim endIdx As Long
    Dim rng As Range

    Set vars = Variables()
    Set design = DesignFormat()

    startIdx = CLng(vars.Value(varName:=firstVarName, colName:="column index"))
    endIdx = CLng(vars.Value(varName:=lastVarName, colName:="column index"))

    With this.wksh
        Set rng = .Range(.Cells(startIdx, VLIST_SEC_COL), .Cells(endIdx, VLIST_SEC_COL))
        rng.Cells(1, 1).Value = sectionName
    End With

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
'@param sheetName String. The current sheet name.
Private Sub FormatSubSection(ByVal subSectionName As String, _
                              ByVal startRow As Long, _
                              ByVal endRow As Long, _
                              ByVal sheetName As String)
    If this.layer = SectionBuilderModeHList Then
        FormatHListSubSection subSectionName, startRow, endRow, sheetName
    Else
        FormatVListSubSection subSectionName, startRow, endRow, sheetName
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
'@param sheetName String. The current sheet name.
Private Sub FormatHListSubSection(ByVal subSectionName As String, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long, _
                                   ByVal sheetName As String)
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim design As ILLFormat
    Dim varRng As Range
    Dim firstVarName As String
    Dim lastVarName As String
    Dim startCol As Long
    Dim endCol As Long
    Dim crfLine As Long
    Dim rng As Range
    Dim sh As Worksheet

    Set dict = Dictionary()
    Set vars = Variables()
    Set design = DesignFormat()
    Set varRng = dict.DataRange("variable name")

    firstVarName = varRng.Cells(startRow, 1).Value
    lastVarName = varRng.Cells(endRow, 1).Value
    startCol = CLng(vars.Value(colName:="column index", varName:=firstVarName))
    endCol = CLng(vars.Value(colName:="column index", varName:=lastVarName))
    crfLine = CLng(vars.Value(colName:="crf index", varName:=firstVarName))

    ' Main sheet subsection
    Set sh = this.wksh
    With sh
        Set rng = .Range(.Cells(HLIST_SUBSEC_ROW, startCol), .Cells(HLIST_SUBSEC_ROW, endCol))
        rng.Cells(1, 1).Value = subSectionName
    End With
    design.ApplyFormat rng, scope:=HListSubSection

    ' Printed sheet subsection
    If Not this.printWksh Is Nothing Then
        Set sh = this.printWksh
        With sh
            Set rng = .Range(.Cells(HLIST_SUBSEC_ROW, startCol), .Cells(HLIST_SUBSEC_ROW, endCol))
            rng.Cells(1, 1).Value = subSectionName
        End With
        design.ApplyFormat rng, scope:=HListSubSection
    End If

    ' CRF sheet subsection
    If (subSectionName <> vbNullString) And (Not this.crfWksh Is Nothing) Then
        Set sh = this.crfWksh
        With sh
            Set rng = .Range(.Cells(crfLine - 3, 1), .Cells(crfLine - 3, 7))
            rng.Cells(1, 1).Value = subSectionName
        End With
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
'@param sheetName String. The current sheet name.
Private Sub FormatVListSubSection(ByVal subSectionName As String, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long, _
                                   ByVal sheetName As String)
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim design As ILLFormat
    Dim varRng As Range
    Dim firstVarName As String
    Dim lastVarName As String
    Dim startIdx As Long
    Dim endIdx As Long
    Dim rng As Range

    Set dict = Dictionary()
    Set vars = Variables()
    Set design = DesignFormat()
    Set varRng = dict.DataRange("variable name")

    firstVarName = varRng.Cells(startRow, 1).Value
    lastVarName = varRng.Cells(endRow, 1).Value
    startIdx = CLng(vars.Value(colName:="column index", varName:=firstVarName))
    endIdx = CLng(vars.Value(colName:="column index", varName:=lastVarName))

    With this.wksh
        Set rng = .Range(.Cells(startIdx, VLIST_SUBSEC_COL), .Cells(endIdx, VLIST_SUBSEC_COL))
        rng.Cells(1, 1).Value = subSectionName
    End With

    design.ApplyFormat rng, scope:=VListSubSection
End Sub


'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise an error with the class name as source
'@param errorMessage String. Human-readable description of the failure.
'@throws VbObjectError+516 Always raises the specified error.
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, "Error : " & errorMessage
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ISectionBuilder contract.

Private Sub ISectionBuilder_Build(ByVal sheetName As String, ByVal startRow As Long)
    Build sheetName, startRow
End Sub
