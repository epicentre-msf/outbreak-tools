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

'@section Factory
'===============================================================================

'@description Create a new SectionBuilder for the given layer and dependencies.
'@param layer SectionBuilderMode enum value (HList=1, VList=2). Use Byte for macOS compatibility.
'@param specs The linelist specifications object.
'@param wksh The target main worksheet.
'@param printWksh Optional printed companion worksheet (HList only).
'@param crfWksh Optional CRF companion worksheet (HList only, may be set later via InitializeCRFSheet).
'@param dropdownObj Optional dropdown lists manager.
'@param customDropdownObj Optional custom dropdown lists manager.
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

Public Property Get Self() As ISectionBuilder
    Set Self = Me
End Property

'@section Internal setters (used by factory only)
'===============================================================================

Public Property Let InternalLayer(ByVal layer As Byte)
    this.layer = layer
End Property

Public Property Set InternalSpecs(ByVal specs As ILinelistSpecs)
    Set this.specs = specs
End Property

Public Property Set InternalWksh(ByVal wksh As Worksheet)
    Set this.wksh = wksh
End Property

Public Property Set InternalPrintWksh(ByVal wksh As Worksheet)
    Set this.printWksh = wksh
End Property

Public Property Set InternalCRFWksh(ByVal wksh As Worksheet)
    Set this.crfWksh = wksh
End Property

Public Property Set InternalDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.dropdown = dropdownObj
End Property

Public Property Set InternalCustomDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.customDropdown = dropdownObj
End Property

'@section Cached accessors
'===============================================================================

Private Property Get Dictionary() As ILLdictionary
    If this.cachedDict Is Nothing Then
        Set this.cachedDict = this.specs.Dictionary()
    End If
    Set Dictionary = this.cachedDict
End Property

Private Property Get Variables() As ILLVariables
    If this.cachedVars Is Nothing Then
        Set this.cachedVars = LLVariables.Create(Dictionary())
    End If
    Set Variables = this.cachedVars
End Property

Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = this.specs.DesignFormat()
End Property

'@section Main entry point
'===============================================================================

'@description Build all sections, subsections, and variables for the given sheet.
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

'@section Section boundary detection
'===============================================================================

'@description Find the end row of a section (last row with same section name and sheet name).
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

'@section Subsection processing
'===============================================================================

'@description Build all subsections within a section range.
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

'@description Find the end row of a subsection.
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

'@section Variable building
'===============================================================================

'@description Build all variables within a subsection range using VarWriter.
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

'@section Section formatting
'===============================================================================

'@description Format a section header by merging cells and applying section format.
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

'@description Format an HList section header (horizontal merge across columns at SECROW).
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

'@description Format a VList section header (vertical merge across rows at SECCOL).
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

'@section Subsection formatting
'===============================================================================

'@description Format a subsection header by merging cells and applying format.
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

'@description Format an HList subsection header (horizontal merge at SUBSECROW).
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

'@description Format a VList subsection header (vertical merge at SUBSECCOL).
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

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, "Error : " & errorMessage
End Sub

'@section Interface implementation
'===============================================================================

Private Sub ISectionBuilder_Build(ByVal sheetName As String, ByVal startRow As Long)
    Build sheetName, startRow
End Sub
