VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@PredeclaredId
'@Folder("Sections")
'@ModuleDescription("Consolidated variable writer for HList, VList and CRF companion sheets")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'Sheet layers a variable is written to. The enum lives in this class, so any
'class that depends on VarWriter resolves it. `test-registry.yml` records the
'same move under `folder: sections`.
Public Enum VarWriterLayer
    VarWriterLayerHList = 1
    VarWriterLayerVList = 2
End Enum


'@class VarWriter
'@description
'Writes a single dictionary variable to one or more linelist worksheets.
'For HList layers the writer targets the main sheet, a printed companion,
'and a CRF companion; for VList layers only the main sheet is written.
'Each invocation of WriteVariable applies label text, number formatting,
'formulas, dropdown choices, data validation, and conditional formatting
'to the resolved cell range.
'
'ONE WRITER PER SHEET
'-------------------------------------------------------------------------------
'The writer is built once per data entry sheet and used for every variable of
'that sheet. `SectionBuilder.Build` owns it. Everything the class caches -- the
'dictionary block, the column map, the hidden name store, the design format,
'the translated messages -- is worth caching only over a whole sheet, and the
'table anchor is set on the first variable the writer is handed.
'
'HOW A DICTIONARY VALUE IS READ
'-------------------------------------------------------------------------------
'The dictionary data block and its header row are read once, as two worksheet
'crossings, and every `ValueOf` after that answers from memory. A column added
'to the dictionary after the map was built is picked up on the first miss, the
'same way `LLVariables.ResolveColumnIndex` re-reads its header snapshot.
'A variable name is matched without regard to case here; `LLdictionary.Prepare`
'makes the names unique, so a name that differs only by case is a fault in the
'sheet.
'@depends LinelistSpecs, LLdictionary, LLVariables, LLSheets, LLFormat,
'TranslationObject, FormulaData, Formulas, DropdownLists, LLGeo, HiddenNames,
'Checking, BetterArray
'@version 1.1 (2026-07-31)


Private Type TVarWriter
    layer As Byte
    specs As LinelistSpecs
    wksh As Worksheet
    hn As HiddenNames
    printWksh As Worksheet
    crfWksh As Worksheet
    dropdown As DropdownLists
    customDropdown As DropdownLists
    varName As String
    cachedRange As Range
    cachedPrintRange As Range
    cachedCRFRange As Range
    cachedDict As LLdictionary
    cachedVars As LLVariables
    cachedSheets As LLSheets
    cachedFormat As LLFormat
    cachedTrans As TranslationObject
    cachedFormulaData As FormulaData
    calculatedMessage As String
    mandatoryMessage As String
    customChoiceMessage As String
    messagesLoaded As Boolean
    dataValues As Variant
    columnOffsets As Collection
    rowOfVariable As Collection
    columnCount As Long
    dataMapLoaded As Boolean
    rowValues As Variant
    rowLoaded As Boolean
    mainAnchorSet As Boolean
    printAnchorSet As Boolean
    scratchCell As Range
    hasChecks As Boolean
    sealed As Boolean
End Type

Private this As TVarWriter
Private Const CLASS_NAME As String = "VarWriter"
Private Const HLIST_START_ROW As Long = 8
Private Const VLIST_START_COL As Long = 5

'Column index the first variable of a sheet carries when the dictionary is
'numbered as `LLdictionary.Prepare` leaves it. It is the fallback for the table
'anchor; the first variable the writer is handed is the primary rule.
Private Const HLIST_ANCHOR_INDEX As Long = 1
Private Const VLIST_ANCHOR_INDEX As Long = 4

'Checking support
Private internalChecks As Checking

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new VarWriter for the given layer and dependencies
'@details
'Factory method that instantiates and configures a VarWriter. The caller
'must supply a valid LinelistSpecs and target Worksheet at minimum.
'Optional parameters allow attaching printed and CRF companion worksheets
'(used only for HList layers) and dropdown list managers.
'@param layer Byte. VarWriterLayer enum value (HList=1, VList=2). Use Byte for macOS compatibility.
'@param specs LinelistSpecs. The linelist specifications object providing dictionary, format, translations, etc.
'@param wksh Worksheet. The target main worksheet.
'@param printWksh Optional Worksheet. Printed companion worksheet (HList only). Defaults to Nothing.
'@param crfWksh Optional Worksheet. CRF companion worksheet (HList only). Defaults to Nothing.
'@param dropdownObj Optional DropdownLists. Dropdown lists manager. Defaults to Nothing.
'@param customDropdownObj Optional DropdownLists. Custom dropdown lists manager (for choice_custom). Defaults to Nothing.
'@return VarWriter. A fully initialised writer bound to the given layer and worksheets.
'@throws ProjectError.ObjectNotInitialized When specs or wksh is Nothing.
'@throws ProjectError.InvalidArgument When layer names neither VarWriterLayer member.
'@export
Public Function Create(ByVal layer As Byte, _
                       ByVal specs As LinelistSpecs, _
                       ByVal wksh As Worksheet, _
                       Optional ByVal printWksh As Worksheet = Nothing, _
                       Optional ByVal crfWksh As Worksheet = Nothing, _
                       Optional ByVal dropdownObj As DropdownLists = Nothing, _
                       Optional ByVal customDropdownObj As DropdownLists = Nothing) As VarWriter

    If specs Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Specifications object is required"
    If wksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Target worksheet is required"

    If (layer <> VarWriterLayerHList) And (layer <> VarWriterLayerVList) Then
        ThrowError ProjectError.InvalidArgument, _
                   "Layer should be VarWriterLayerHList or VarWriterLayerVList. Found: " & layer
    End If

    Dim newWriter As VarWriter
    Set newWriter = New VarWriter
    Set newWriter.InternalSpecs = specs
    Set newWriter.InternalWksh = wksh
    newWriter.InternalLayer = layer
    If Not printWksh Is Nothing Then Set newWriter.InternalPrintWksh = printWksh
    If Not crfWksh Is Nothing Then Set newWriter.InternalCRFWksh = crfWksh
    If Not dropdownObj Is Nothing Then Set newWriter.InternalDropdown = dropdownObj
    If Not customDropdownObj Is Nothing Then Set newWriter.InternalCustomDropdown = customDropdownObj
    newWriter.Seal
    Set Create = newWriter
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
'@description Properties used by the factory only during instantiation.
'===============================================================================

'@label:internal-specs
'@prop-title Assign the linelist specifications object
Public Property Set InternalSpecs(ByVal specs As LinelistSpecs)
    GuardNotSealed "InternalSpecs"
    Set this.specs = specs
End Property

'@label:internal-wksh
'@prop-title Assign the target main worksheet
Public Property Set InternalWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalWksh"
    Set this.wksh = wksh
    Set this.hn = HiddenNames.Create(this.wksh)
End Property

'@label:internal-layer
'@prop-title Assign the writer layer (HList or VList)
Public Property Let InternalLayer(ByVal layer As Byte)
    GuardNotSealed "InternalLayer"
    this.layer = layer
End Property

'@label:internal-print-wksh
'@prop-title Assign the printed companion worksheet
Public Property Set InternalPrintWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalPrintWksh"
    Set this.printWksh = wksh
End Property

'@label:internal-crf-wksh
'@prop-title Assign the CRF companion worksheet
Public Property Set InternalCRFWksh(ByVal wksh As Worksheet)
    GuardNotSealed "InternalCRFWksh"
    Set this.crfWksh = wksh
End Property

'@label:internal-dropdown
'@prop-title Assign the dropdown lists manager
Public Property Set InternalDropdown(ByVal dropdownObj As DropdownLists)
    GuardNotSealed "InternalDropdown"
    Set this.dropdown = dropdownObj
End Property

'@label:internal-custom-dropdown
'@prop-title Assign the custom dropdown lists manager
Public Property Set InternalCustomDropdown(ByVal dropdownObj As DropdownLists)
    GuardNotSealed "InternalCustomDropdown"
    Set this.customDropdown = dropdownObj
End Property

'@section Cached Accessors
'@description
'Every collaborator the writer uses is resolved once and held. Each of these is
'reached several times per variable, so a pass-through to `this.specs` on every
'use was paid a few hundred times per sheet.
'===============================================================================

'@label:dictionary
'@prop-title Return the cached dictionary object
'@return LLdictionary. The linelist dictionary.
Private Property Get Dictionary() As LLdictionary
    If this.cachedDict Is Nothing Then
        Set this.cachedDict = this.specs.Dictionary()
    End If
    Set Dictionary = this.cachedDict
End Property

'@label:variables
'@prop-title Return the cached variables helper
'@details
'Kept for `Contains`, which the two conditional formatting routines ask for a
'variable that is not the one being written. Reading a column of the CURRENT
'variable goes through `ValueOf`, which answers from the cached dictionary
'block.
'@return LLVariables. The variables helper object.
Private Property Get Variables() As LLVariables
    If this.cachedVars Is Nothing Then
        Set this.cachedVars = LLVariables.Create(Dictionary())
    End If
    Set Variables = this.cachedVars
End Property

'@label:sheets-object
'@prop-title Return the cached sheets helper
'@details
'Provides VariableAddress and SheetInfo lookups used when building
'conditional-formatting formulas and resolving table anchors.
'@return LLSheets. The sheets helper object.
Private Property Get SheetsObject() As LLSheets
    If this.cachedSheets Is Nothing Then
        Set this.cachedSheets = LLSheets.Create(Dictionary())
    End If
    Set SheetsObject = this.cachedSheets
End Property

'@label:design-format
'@prop-title Return the design format object from specs
'@details
'`AddLabel`, `AddFormula`, `AddChoices`, `AddCRFLabel`, `AddCRFType` and
'`AddCRFChoices` each ask for it, so it is resolved once per writer.
'@return LLFormat. The design format provider.
Private Property Get DesignFormat() As LLFormat
    If this.cachedFormat Is Nothing Then
        Set this.cachedFormat = this.specs.DesignFormat()
    End If
    Set DesignFormat = this.cachedFormat
End Property

'@label:trans
'@prop-title Return the translation object from specs
'@return TranslationObject. The translation provider.
Private Property Get Trans() As TranslationObject
    If this.cachedTrans Is Nothing Then
        Set this.cachedTrans = this.specs.TransObject()
    End If
    Set Trans = this.cachedTrans
End Property

'@label:formula-data-obj
'@prop-title Return the formula data object from specs
'@details
'Provides the lookup tables and metadata needed by the Formulas parser.
'@return FormulaData. The formula data provider.
Private Property Get FormulaDataObj() As FormulaData
    If this.cachedFormulaData Is Nothing Then
        Set this.cachedFormulaData = this.specs.FormulaDataObject()
    End If
    Set FormulaDataObj = this.cachedFormulaData
End Property

'@label:ensure-messages
'@sub-title Read the three translated label messages once
'@details
'`MSG_Calculated`, `MSG_Mandatory` and `MSG_CustomChoice` are constant for a
'whole build, and `AddLabel` used to fetch all three per variable and twice per
'HList variable.
Private Sub EnsureMessages()
    Dim transObj As TranslationObject

    If this.messagesLoaded Then Exit Sub

    Set transObj = Trans()
    this.calculatedMessage = transObj.TranslatedValue("MSG_Calculated")
    this.mandatoryMessage = transObj.TranslatedValue("MSG_Mandatory")
    this.customChoiceMessage = transObj.TranslatedValue("MSG_CustomChoice")
    this.messagesLoaded = True
End Sub

'@section Dictionary Row Cache
'===============================================================================
'@description
'The dictionary block and its header row are read once per writer. Every
'`ValueOf` after that is an array read.

'@label:ensure-data-map
'@sub-title Read the dictionary block and build the two lookups
'@details
'Two crossings: the data block and the header row. From them come a header to
'column-offset map and a variable name to row-offset map. Both keep the first
'entry when a name repeats, which is the rule `LLVariables.HeaderColumn` and
'`LLVariables.RowIndex` already apply.
Private Sub EnsureDataMap()
    If this.dataMapLoaded Then Exit Sub
    RebuildDataMap
End Sub

'@label:rebuild-data-map
'@sub-title Read the dictionary block again
'@details
'Called on the first use and again whenever a column name misses, which is how
'a column added to the dictionary after the writer was built resolves.
Private Sub RebuildDataMap()
    Dim dataSh As DataSheet
    Dim headerValues As Variant
    Dim nameOffset As Long
    Dim counter As Long
    Dim headerText As String
    Dim variableText As String

    Set dataSh = Dictionary().Data

    this.dataValues = BlockValues(dataSh.DataRange())
    headerValues = RowOfValues(dataSh.HeaderRange())

    Set this.columnOffsets = New Collection
    this.columnCount = 0
    If Not IsEmpty(headerValues) Then
        this.columnCount = UBound(headerValues)
        For counter = LBound(headerValues) To UBound(headerValues)
            headerText = CellText(headerValues(counter))
            If LenB(headerText) > 0 Then
                'A repeated header keeps its leftmost column. Collection.Add is
                'the only statement under this guard and a duplicate key is the
                'only error it can raise here.
                On Error Resume Next
                    this.columnOffsets.Add counter, headerText
                On Error GoTo 0
            End If
        Next counter
    End If

    Set this.rowOfVariable = New Collection
    nameOffset = ColumnOffset("variable name")

    If (nameOffset > 0) And (Not IsEmpty(this.dataValues)) Then
        For counter = LBound(this.dataValues, 1) To UBound(this.dataValues, 1)
            variableText = CellText(this.dataValues(counter, nameOffset))
            If LenB(variableText) > 0 Then
                'A repeated variable name keeps its first row.
                On Error Resume Next
                    this.rowOfVariable.Add counter, variableText
                On Error GoTo 0
            End If
        Next counter
    End If

    this.dataMapLoaded = True
    this.rowLoaded = False
End Sub

'@label:dictionary-grew-a-column
'@sub-title Test whether the dictionary is wider than the map was built from
'@details
'One property read against a whole block read. `AddLabel` asks for `list auto`
'on every variable, and a dictionary that has not been through
'`LinelistSpecs.Prepare` carries no such column, so a rebuild on every miss made
'each variable pay for a full re-read of the dictionary and two lookup rebuilds
'for an answer that was always empty.
'@return Boolean. True when the header row is wider than it was.
Private Function DictionaryGrewAColumn() As Boolean
    Dim dataSh As DataSheet
    Dim currentCount As Long

    Set dataSh = Dictionary().Data
    currentCount = dataSh.DataEndColumn() - dataSh.DataStartColumn() + 1

    DictionaryGrewAColumn = (currentCount > this.columnCount)
End Function

'@label:column-offset
'@sub-title Offset of a dictionary column inside a data row
'@details
'Whole string, without regard to case, which is what `LLVariables.HeaderColumn`
'answers. A Collection raises on a key it does not hold, and that is how a miss
'is read.
'@param colName String. The dictionary column header.
'@return Long. The 1-based offset, or 0 when the column is absent.
Private Function ColumnOffset(ByVal colName As String) As Long
    If this.columnOffsets Is Nothing Then Exit Function
    If LenB(colName) = 0 Then Exit Function

    On Error Resume Next
        ColumnOffset = this.columnOffsets.Item(colName)
    On Error GoTo 0
End Function

'@label:row-offset
'@sub-title Row of one variable inside the dictionary data block
'@param varName String. The variable name.
'@return Long. The 1-based row, or 0 when the variable is absent.
Private Function RowOffset(ByVal varName As String) As Long
    If this.rowOfVariable Is Nothing Then Exit Function
    If LenB(varName) = 0 Then Exit Function

    On Error Resume Next
        RowOffset = this.rowOfVariable.Item(varName)
    On Error GoTo 0
End Function

'@label:ensure-row-values
'@sub-title Hold the current variable's dictionary row in memory
Private Sub EnsureRowValues()
    Dim rowNumber As Long
    Dim counter As Long
    Dim rowLine() As Variant

    If this.rowLoaded Then Exit Sub

    EnsureDataMap

    this.rowValues = Empty
    this.rowLoaded = True

    rowNumber = RowOffset(this.varName)
    If rowNumber = 0 Then Exit Sub
    If IsEmpty(this.dataValues) Then Exit Sub

    ReDim rowLine(1 To UBound(this.dataValues, 2))
    For counter = 1 To UBound(this.dataValues, 2)
        rowLine(counter) = this.dataValues(rowNumber, counter)
    Next counter

    this.rowValues = rowLine
End Sub

'@label:block-values
'@sub-title Read a Range as a two-dimensional 1-based array
'@details
'`Range.Value` hands back a plain value for one cell and a two-dimensional
'array otherwise. This flattens both cases so callers see one shape.
'@param blockRng Range. The block to read.
'@return Variant. A 1-based two-dimensional array of cell values.
Private Function BlockValues(ByVal blockRng As Range) As Variant
    Dim singleCell() As Variant

    If blockRng Is Nothing Then Exit Function

    If blockRng.Cells.CountLarge = 1 Then
        ReDim singleCell(1 To 1, 1 To 1)
        singleCell(1, 1) = blockRng.Value
        BlockValues = singleCell
        Exit Function
    End If

    BlockValues = blockRng.Value
End Function

'@label:row-of-values
'@sub-title Read a one-row Range as a 1-based list of values
'@param rowRng Range. A single-row block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function RowOfValues(ByVal rowRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long

    If rowRng Is Nothing Then Exit Function

    If rowRng.Cells.CountLarge = 1 Then
        ReDim flat(1 To 1)
        flat(1) = rowRng.Value
        RowOfValues = flat
        Exit Function
    End If

    rawValues = rowRng.Value
    ReDim flat(1 To UBound(rawValues, 2))
    For counter = 1 To UBound(rawValues, 2)
        flat(counter) = rawValues(1, counter)
    Next counter

    RowOfValues = flat
End Function

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'A dictionary cell can hold an error value or a Null, and `CStr` raises on
'both. Every read of the cached block goes through here, the same way every
'read in `LLVariables` goes through its own copy.
'@param cellValue Variant. The raw cell value.
'@return String. The text, or an empty string when there is none.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@section Value Lookup
'===============================================================================

'@label:value-of
'@prop-title Read a dictionary column value for the current variable
'@details
'The class's read hook, and the way every other method retrieves variable
'metadata such as labels, types, formats and control information. An unknown
'column and an unknown variable both answer an empty string, which is what
'`LLVariables.Value` answers for the same two cases.
'@param colName String. The dictionary column name to look up.
'@return String. The column value for the current variable.
Public Property Get ValueOf(ByVal colName As String) As String
    Dim offset As Long

    If LenB(this.varName) = 0 Then Exit Property

    EnsureDataMap
    EnsureRowValues

    offset = ColumnOffset(colName)

    If offset = 0 Then
        'A column the dictionary gained after the map was built, which is what
        'LinelistSpecs.AddListAuto does. The header row is measured first,
        'because a name the dictionary has never carried is asked for once per
        'variable and re-reading the whole block for it would be paid every time.
        If Not DictionaryGrewAColumn() Then Exit Property

        RebuildDataMap
        EnsureRowValues
        offset = ColumnOffset(colName)
        If offset = 0 Then Exit Property
    End If

    If IsEmpty(this.rowValues) Then Exit Property
    If offset > UBound(this.rowValues) Then Exit Property

    ValueOf = CellText(this.rowValues(offset))
End Property

'@label:number-of
'@sub-title Read a dictionary column value as a whole number
'@details
'`column index` and `crf index` were read with a bare `CLng(ValueOf(...))` at
'eight sites, guarded at three of them and unguarded at the rest. This is the
'one rule: a value that is empty or is not a number answers 0, and every caller
'tests for 0.
'@param colName String. The dictionary column name to look up.
'@return Long. The stored number, or 0.
Private Function NumberOf(ByVal colName As String) As Long
    Dim storedText As String

    storedText = ValueOf(colName)
    If LenB(storedText) = 0 Then Exit Function
    If Not IsNumeric(storedText) Then Exit Function

    NumberOf = CLng(storedText)
End Function

'@section Cell Range Resolution
'===============================================================================

'@label:var-range
'@prop-title Resolve the main cell range for the current variable
'@details
'Computes the target cell from the column index dictionary value and the
'layer-specific start position. For HList the cell is in a column at
'HLIST_START_ROW + 1; for VList it is in a row at VLIST_START_COL. The
'result is cached until WriteVariable resets it for the next variable.
'@return Range. The main cell range, or Nothing when the index is missing.
Private Property Get VarRange() As Range
    If this.cachedRange Is Nothing Then
        Dim idx As Long
        idx = NumberOf("column index")
        If idx <= 0 Then Exit Property

        Select Case this.layer
        Case VarWriterLayerHList
            Set this.cachedRange = this.wksh.Cells(HLIST_START_ROW + 1, idx)
        Case VarWriterLayerVList
            Set this.cachedRange = this.wksh.Cells(idx, VLIST_START_COL)
        End Select
    End If
    Set VarRange = this.cachedRange
End Property

'@label:var-print-range
'@prop-title Resolve the printed companion cell range (HList only)
'@details
'Returns the cell on the printed companion worksheet that corresponds to
'the current variable. Uses the same column index as VarRange but targets
'the printWksh. Returns Nothing when no printed worksheet is attached.
'@return Range. The printed companion cell range, or Nothing.
Private Property Get VarPrintRange() As Range
    If this.cachedPrintRange Is Nothing Then
        If Not this.printWksh Is Nothing Then
            Dim idx As Long
            idx = NumberOf("column index")
            If idx <= 0 Then Exit Property
            Set this.cachedPrintRange = this.printWksh.Cells(HLIST_START_ROW + 1, idx)
        End If
    End If
    Set VarPrintRange = this.cachedPrintRange
End Property

'@label:var-crf-range
'@prop-title Resolve the CRF companion cell range (HList only)
'@details
'Returns the cell on the CRF companion worksheet at the row specified by
'the "crf index" dictionary value. Returns Nothing when no CRF worksheet
'is attached or when the crf index is zero or missing.
'@return Range. The CRF companion cell range, or Nothing.
Private Property Get VarCRFRange() As Range
    If this.cachedCRFRange Is Nothing Then
        If Not this.crfWksh Is Nothing Then
            Dim crfLine As Long
            crfLine = NumberOf("crf index")

            If crfLine > 0 Then
                Set this.cachedCRFRange = this.crfWksh.Cells(crfLine, 1)
            End If
        End If
    End If
    Set VarCRFRange = this.cachedCRFRange
End Property

'@section Entry Point
'===============================================================================

'@label:write-variable
'@sub-title Write one variable to the target worksheet(s)
'@details
'Main public method called once per variable. It stores the variable name,
'resets the cached row and cell ranges, then dispatches to WriteMainSheet. For
'HList layers it additionally dispatches to WritePrintSheet and WriteCRFSheet
'for the companion worksheets. A variable the dictionary gives no column index
'is skipped and filed through Checking, because the cell it would be written to
'cannot be resolved.
'@param varName String. The variable name as it appears in the dictionary.
'@export
Public Sub WriteVariable(ByVal varName As String)
    this.varName = varName

    ' Reset the cached row and cell ranges for the new variable
    this.rowLoaded = False
    Set this.cachedRange = Nothing
    Set this.cachedPrintRange = Nothing
    Set this.cachedCRFRange = Nothing

    If VarRange() Is Nothing Then
        LogCheck "Variable " & varName & " carries no column index, so it was not written", _
                 checkingError
        Exit Sub
    End If

    ' Write to main sheet
    WriteMainSheet

    ' HList: also write to printed and CRF companions
    If this.layer = VarWriterLayerHList Then
        WritePrintSheet
        WriteCRFSheet
    End If
End Sub

'@section Sheet Dispatch
'===============================================================================

'@label:write-main-sheet
'@sub-title Dispatch all writing steps for the main worksheet
'@details
'Orchestrates the sequence of writing operations on the main worksheet:
'label, formula, type formatting, choices, validation, and conditional
'formatting. The exact set of conditional formatting steps depends on
'the active layer (HList applies geo, standard, and unique formatting;
'VList applies standard formatting only).
Private Sub WriteMainSheet()
    AddLabel
    AddFormula
    AddType
    AddChoices

    On Error Resume Next
    AddValidation
    On Error GoTo 0

    ' HList-specific conditional formatting
    If this.layer = VarWriterLayerHList Then

        On Error Resume Next
        AddGeoConditionalFormatting
        On Error GoTo 0

        AddConditionalFormatting
        AddUniqueConditionalFormatting
    End If

    ' VList also has conditional formatting
    If this.layer = VarWriterLayerVList Then
        AddConditionalFormatting
    End If
End Sub

'@label:write-print-sheet
'@sub-title Write variable label to the printed companion sheet (HList only)
'@details
'Delegates to AddLabel with usePrinted=True so that the label is written
'to the printed companion worksheet instead of the main sheet. Exits
'silently when no printed worksheet is attached.
Private Sub WritePrintSheet()
    If this.printWksh Is Nothing Then Exit Sub
    AddLabel usePrinted:=True
End Sub

'@label:write-crf-sheet
'@sub-title Write variable data to the CRF companion sheet (HList only)
'@details
'Resolves the CRF cell range and delegates to AddCRFLabel, AddCRFType,
'and AddCRFChoices. Exits silently when no CRF worksheet is attached or
'when the variable has no CRF index.
Private Sub WriteCRFSheet()
    If this.crfWksh Is Nothing Then Exit Sub
    Dim crfRng As Range
    Set crfRng = VarCRFRange()
    If crfRng Is Nothing Then Exit Sub

    AddCRFLabel crfRng
    AddCRFType crfRng
    AddCRFChoices crfRng
End Sub

'@section Label Writing
'===============================================================================

'@label:add-label
'@sub-title Write variable label, sublabel, name, notes, and formatting
'@details
'Writes the main label, sublabel, variable name, notes, control info, and
'design formatting to the target cell range. For HList layers this includes
'header cells, named ranges for table anchors, comments for notes, hidden
'control rows, geo-specific formatting, editable labels, and column hiding.
'For VList layers the layout is row-based with adjacent cells for control
'info and label formatting.
'@param usePrinted Optional Boolean. When True, writes to the printed companion range instead of the main range. Defaults to False.
Private Sub AddLabel(Optional ByVal usePrinted As Boolean = False)
    Dim varMainLab As String
    Dim varName As String
    Dim varNote As String
    Dim varSubLab As String
    Dim varStatus As String
    Dim varListAutoOrigin As String
    Dim varControl As String
    Dim varIndex As Long
    Dim cellRng As Range
    Dim mainLabLength As Long
    Dim design As LLFormat
    Dim varToPrint As String
    Dim varEditLab As String

    EnsureMessages
    Set design = DesignFormat()

    varMainLab = ValueOf("main label")
    varName = ValueOf("variable name")
    varNote = ValueOf("note")
    varSubLab = ValueOf("sub label")
    varControl = ValueOf("control")
    varEditLab = ValueOf("editable label")
    varStatus = ValueOf("status")
    varListAutoOrigin = ValueOf("list auto")
    varIndex = NumberOf("column index")

    ' Resolve the target cell range based on layer and printed flag
    If this.layer = VarWriterLayerHList Then
        If usePrinted Then
            Set cellRng = VarPrintRange()
        Else
            Set cellRng = VarRange()
        End If
    Else
        ' VList
        Set cellRng = VarRange()
    End If

    If cellRng Is Nothing Then Exit Sub

    ' Sublabel enrichment
    If (varControl = "formula") Or (varControl = "case_when") Or _
       (varControl = "choice_formula") Then
        varSubLab = varSubLab & Chr(10) & this.calculatedMessage
    End If

    If (varControl = "choice_custom") Then
        varSubLab = varSubLab & Chr(10) & "(" & this.customChoiceMessage & ")"
    End If

    mainLabLength = Len(varMainLab)
    If (varSubLab <> vbNullString) Then
        varMainLab = varMainLab & Chr(10) & varSubLab
    End If

    ' Status to add to notes
    If varStatus = "mandatory" Then varNote = this.mandatoryMessage & Chr(10) & varNote

    ' --- HList-specific label writing ---
    If this.layer = VarWriterLayerHList Then
        ' Table start anchor: the first variable the writer is handed carries
        ' it. The canonical column index stays as a fallback, so a writer given
        ' the variables out of order still leaves the anchor on the right cell.
        If Not usePrinted Then
            If (Not this.mainAnchorSet) Or (varIndex = HLIST_ANCHOR_INDEX) Then
                cellRng.Offset(-1).Name = ValueOf("table name") & "_START"
                this.mainAnchorSet = True
            End If
        Else
            If (Not this.printAnchorSet) Or (varIndex = HLIST_ANCHOR_INDEX) Then
                cellRng.Offset(-1).Name = ValueOf("table name") & "_PRINTSTART"
                this.printAnchorSet = True
            End If
        End If

        ' Main label and sub label
        If (varControl <> "choice_custom") Or usePrinted Then
            cellRng.Offset(-2).Value = varMainLab
            design.ApplyFormat cellRng.Offset(-2).Characters(Start:=1, Length:=mainLabLength), _
                scope:=HListMainLab
            design.ApplyFormat cellRng.Offset(-2).Characters(Start:=mainLabLength + 1), _
                scope:=HListSublab
        End If

        ' Variable name
        cellRng.Offset(-1).Value = varName
        If Not usePrinted Then cellRng.Offset(-2).Name = varName

        ' Variable notes. The comment is cleared first, because Excel raises
        ' 1004 on AddComment when the cell already carries one and a rebuild
        ' over an existing sheet reaches this twice.
        If varNote <> vbNullString And (Not usePrinted) Then
            cellRng.Offset(-2).ClearComments
            cellRng.Offset(-2).AddComment
            cellRng.Offset(-2).Comment.Text Text:=varNote
            cellRng.Offset(-2).Comment.Visible = False
        End If

        ' Control info in hidden rows - Use HiddenNames instead
        this.hn.EnsureName varName & " -- " & "control", varControl, HiddenNameTypeString
        this.hn.EnsureName varName & " -- " & "listauto", varListAutoOrigin, HiddenNameTypeString
        this.hn.EnsureName varName & " -- " & "address", cellRng.Address, HiddenNameTypeString

        ' Sheet-level flag for listauto presence (used by EventLinelist)
        If varListAutoOrigin <> vbNullString Then
            this.hn.EnsureName "has_listauto", "yes", HiddenNameTypeString
        End If

        ' Geo labels
        If (InStr(1, varControl, "geo") > 0) Or (varControl = "hf") Or _
            (InStr(1, varName, "pcode_adm") > 0) Then
            design.ApplyFormat cellRng.Offset(-2), scope:=HListGeoHeader
            design.ApplyFormat cellRng, scope:=HListGeo
        End If

        ' Editable labels
        'Add a tag for editable
        If ((varEditLab = "yes") Or (varControl = "custom")) Then
            this.hn.EnsureName varName & " -- " & "editable", "yes", HiddenNameTypeString
            cellRng.Offset(-2).Locked = False
        End If

        ' Hidden columns
        varToPrint = ValueOf("register book")
        If ((varStatus = "hidden") Or (varStatus = "optional, hidden")) And _
           (Not usePrinted) And (cellRng.Column > 1) Then
            cellRng.EntireColumn.Hidden = True
        ElseIf ((varToPrint = "hidden" Or varStatus = "hidden") And usePrinted) Then
            cellRng.EntireColumn.Hidden = True
        ElseIf (usePrinted And (varToPrint = "print, vertical header")) Then
            cellRng.Offset(-2).Orientation = 90
            cellRng.Offset(-2).EntireColumn.AutoFit
        End If

        ' Header cell format
        design.ApplyFormat cellRng.Offset(-2), scope:=HListSimpleHeaderCell
        cellRng.Locked = False

    ' --- VList-specific label writing ---
    Else
        ' Table start anchor, on the same rule as the HList branch above.
        If (Not this.mainAnchorSet) Or (varIndex = VLIST_ANCHOR_INDEX) Then
            cellRng.Name = ValueOf("table name") & "_START"
            this.mainAnchorSet = True
        End If

        ' Main label and sub label
        cellRng.Offset(, -1).Value = varMainLab

        If varControl <> "choice_custom" Then
            design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=1, Length:=mainLabLength), _
                scope:=VListMainLab
            design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=mainLabLength + 1), _
                scope:=VListSublab
        End If

        ' Variable name
        this.hn.EnsureName varName & " -- " & "address", cellRng.Address, HiddenNameTypeString
        cellRng.Name = varName

        ' Variable note
        If varNote <> vbNullString Then
            cellRng.ClearComments
            cellRng.AddComment
            cellRng.Comment.Text Text:=varNote
            cellRng.Comment.Visible = False
        End If

        ' Hidden row
        If (varStatus = "hidden") Or (varStatus = "optional, hidden") Then
            cellRng.EntireRow.Hidden = True
        End If

        ' Editable labels
        If ((varEditLab = "yes") Or (varControl = "custom")) Then
            this.hn.EnsureName varName & " -- " & "editable", "yes", HiddenNameTypeString
            cellRng.Offset(, -1).Locked = False
        End If

        cellRng.Locked = False
        cellRng.Offset(, -2).Locked = True
        cellRng.Offset(, -3).Locked = True

    End If
End Sub

'@section Type Formatting
'===============================================================================

'@label:add-type
'@sub-title Apply Excel number format based on variable type and format
'@details
'Reads the variable type (text, date, decimal, integer) and optional
'format string from the dictionary, then sets the NumberFormat property
'on the main cell range. Supports built-in formats such as round, percentage,
'euros, and dollars as well as custom format strings. A numeric variable whose
'format reads `text` keeps the number format it already has: Excel refuses an
'empty NumberFormat, and this method sits outside every suppression in
'WriteMainSheet, so assigning one used to stop the sheet.
Private Sub AddType()
    Dim varTypes As String
    Dim varFormat As String
    Dim cellRng As Range
    Dim nbDecimals As Integer
    Dim zerosDecimals As String
    Dim formatText As String

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    varTypes = LCase(Trim$(ValueOf("variable type")))
    varFormat = LCase(Trim$(ValueOf("variable format")))
    nbDecimals = 1

    Select Case varTypes

    Case "text"
        cellRng.NumberFormat = "@"

    Case "date"
        cellRng.NumberFormat = "d-mmm-yyyy"
        On Error Resume Next
        If varFormat <> vbNullString Then cellRng.NumberFormat = varFormat
        On Error GoTo 0

    Case "decimal", "integer"
        If varFormat = vbNullString Then
            formatText = IIf(varTypes = "decimal", "#,##0.00;-#,##0.00;0.00;@", "###0;-###0;0;@")
        ElseIf varFormat = "integer" Then
            formatText = "###0;-###0;0;@"
        ElseIf varFormat = "round0" Then
            formatText = "#,##0;-#,##0;0;@"
        ElseIf varFormat = "percentage0" Then
            formatText = "0%"
        ElseIf (InStr(1, varFormat, "round") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "round", vbNullString))
            On Error GoTo 0
            zerosDecimals = String(nbDecimals, "0")
            formatText = "#,###0." & zerosDecimals & ";-#,###0." & zerosDecimals & ";0." & zerosDecimals & ";@"
        ElseIf (InStr(1, varFormat, "percentage") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "percentage", vbNullString))
            On Error GoTo 0
            zerosDecimals = String(nbDecimals, "0")
            formatText = "0." & zerosDecimals & "%"
        ElseIf varFormat = "euros" Then
            'ChrW reads 8364 as a Unicode code point on every machine. Chr is
            'documented for 0 to 255 and reads its argument as a byte position
            'in the machine's ANSI codepage.
            formatText = "#,##0.00 [$" & ChrW$(8364) & "-x-euro1];-#,##0.00 [$" & ChrW$(8364) & _
                         "-x-euro1];0.00 [$" & ChrW$(8364) & "-x-euro1];@"
        ElseIf varFormat = "dollars" Then
            formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"
        ElseIf varFormat <> "text" Then
            formatText = "custom"
        End If

        If formatText = vbNullString Then Exit Sub

        If formatText <> "custom" Then
            cellRng.NumberFormat = formatText
        Else
            On Error Resume Next
            cellRng.NumberFormat = varFormat
            On Error GoTo 0
        End If

    End Select
End Sub

'@section Formula Application
'===============================================================================

'@label:add-formula
'@sub-title Parse and apply a formula to the variable cell
'@details
'Handles formula, case_when, and choice_formula control types. Parses the
'control details string through the Formulas class, validates the result,
'and writes the Excel formula to the cell.
'
'A formula the parser rejects leaves the cell alone. The calculated-cell
'styling and the lock both sit inside the same guard, so a cell that reads as
'read-only is read-only, and the reason the parser gave is filed through
'Checking. Before this the styling was applied whatever the parser answered,
'while `AddLabel` had already unlocked the cell -- so the user saw a greyed
'calculated cell they could type into and the value went into the export.
Private Sub AddFormula()
    Dim varForm As String
    Dim excelForm As String
    Dim cellRng As Range
    Dim varFormData As FormulaData
    Dim dict As LLdictionary
    Dim varFormObject As Formulas
    Dim design As LLFormat
    Dim varControl As String

    varControl = ValueOf("control")
    If (varControl <> "formula") And (varControl <> "case_when") And _
       (varControl <> "choice_formula") Then Exit Sub

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    Set dict = Dictionary()
    Set varFormData = FormulaDataObj()
    Set design = DesignFormat()
    varForm = ValueOf("control details")

    Set varFormObject = Formulas.Create(dict, varFormData, varForm)

    If varFormObject.Valid(formulaType:="linelist") Then
        excelForm = varFormObject.ParsedLinelistFormula()
        If FailedFormula(excelForm) Then
            excelForm = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
        End If
        cellRng.formula = "= " & excelForm
        cellRng.Locked = True

        ' Apply calculated formula format
        design.ApplyFormat cellRng, scope:=HListCalculatedFormulaCell
        If this.layer = VarWriterLayerHList Then
            design.ApplyFormat cellRng.Offset(-2), scope:=HListCalculatedFormulaHeader
        End If
    Else
        LogCheck "The formula of variable " & this.varName & " was rejected: " & varForm, _
                 checkingError
    End If

    ' Harvest formula parsing checkings
    If varFormObject.HasChecking Then
        HarvestCollaboratorCheckings varFormObject.CheckingValues
    End If
End Sub

'@section Choice Setup
'===============================================================================

'@label:add-choices
'@sub-title Set up dropdown validation for choice, list, and geo variables
'@details
'Handles choice_manual, choice_multiple, choice_custom, list_auto, and geo1
'control types. For each type it populates a BetterArray of categories,
'registers them with the appropriate DropdownLists manager, and calls
'SetValidation to apply dropdown data validation to the cell. The
'choice_custom branch additionally creates a hyperlink from the label cell
'to the dropdown worksheet and adds a return link. The geo1 branch sets up
'admin1 through admin4 dropdown cascades.
'
'The two dropdown managers are optional at creation, so a writer built without
'the one a control type needs files a checking and leaves the variable without
'its dropdown. The same holds for a geo variable in a linelist with no
'geography.
Private Sub AddChoices()
    Dim varControl As String
    Dim dropName As String
    Dim varAlert As String
    Dim varMessage As String
    Dim cat As BetterArray
    Dim drop As DropdownLists
    Dim cellRng As Range
    Dim geo As LLGeo
    Dim varMainLab As String
    Dim varSubLab As String
    Dim sh As Worksheet
    Dim labRng As Range
    Dim targetAddress As String
    Dim mainLabLength As Long
    Dim design As LLFormat

    varControl = ValueOf("control")

    'A control can carry its argument in brackets, as in choice_manual(list).
    'The bracket is tested for rather than suppressing the split: VBA answers
    'Split(vbNullString, "(") with a ZERO-LENGTH array, so reading element 0 of
    'it raises subscript out of range -- and a variable with no control at all
    'is the common case here.
    If InStr(1, varControl, "(") > 0 Then varControl = Split(varControl, "(")(0)

    ' Determine which controls apply based on layer
    If this.layer = VarWriterLayerHList Then
        If (varControl <> "choice_manual") And (varControl <> "list_auto") And _
           (varControl <> "geo1") And (varControl <> "choice_custom") And _
           (varControl <> "choice_multiple") Then Exit Sub
    Else
        ' VList: no list_auto, no geo1
        If (varControl <> "choice_manual") And _
           (varControl <> "choice_custom") And (varControl <> "choice_multiple") Then Exit Sub
    End If

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub
    Set sh = cellRng.Parent

    If varControl = "choice_custom" Then
        Set drop = this.customDropdown
    Else
        Set drop = this.dropdown
    End If

    If drop Is Nothing Then
        LogCheck "Variable " & this.varName & " is a " & varControl & _
                 " and the writer was given no dropdown manager for it, so it has no dropdown", _
                 checkingError
        Exit Sub
    End If

    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")

    Select Case varControl

    Case "choice_manual", "choice_multiple"
        dropName = this.varName
        Set cat = this.specs.Categories(dropName)
        drop.Add cat, dropName

    Case "choice_custom"
        EnsureMessages
        Set design = DesignFormat()
        varSubLab = ValueOf("sub label")
        varMainLab = ValueOf("main label")
        mainLabLength = Len(varMainLab)

        dropName = this.varName
        Set cat = this.specs.Categories(dropName)
        If cat.Length = 0 Then cat.Push vbNullString, vbNullString

        drop.Add cat, dropName, addLabel:=True, counterPrefix:=this.customChoiceMessage

        Set labRng = drop.LabelRange(dropName)
        targetAddress = "'" & labRng.Parent.Name & "'" & "!" & labRng.Address

        ' Add hyperlink to the label cell
        Dim anchorRng As Range
        If this.layer = VarWriterLayerHList Then
            Set anchorRng = cellRng.Offset(-2)
        Else
            Set anchorRng = cellRng.Offset(, -1)
        End If

        sh.Hyperlinks.Add _
            Anchor:=anchorRng, _
            Address:="", _
            SubAddress:=targetAddress, _
            TextToDisplay:=""

        anchorRng.Font.UnderLine = xlUnderlineStyleNone

        ' Rebuild label with custom choice counter
        varSubLab = varSubLab & Chr(10) & "(" & labRng.Value & ")"
        varMainLab = varMainLab & Chr(10) & varSubLab

        anchorRng.Value = varMainLab

        ' Re-apply label formatting
        If this.layer = VarWriterLayerHList Then
            design.ApplyFormat anchorRng.Characters(Start:=1, Length:=mainLabLength), _
                scope:=HListMainLab
            design.ApplyFormat anchorRng.Characters(Start:=mainLabLength + 1), _
                scope:=HListSublab
        Else
            design.ApplyFormat anchorRng.Characters(Start:=1, Length:=mainLabLength), _
                scope:=VListMainLab
            design.ApplyFormat anchorRng.Characters(Start:=mainLabLength + 1), _
                scope:=VListSublab
        End If

        ' Add return link from the dropdown worksheet
        drop.AddReturnLink listName:=dropName, returnRng:=anchorRng

    Case "list_auto"
        ' HList only
        dropName = ValueOf("control details")
        Set cat = New BetterArray
        cat.Push ""
        drop.Add cat, dropName

    Case "geo1"
        ' HList only
        dropName = "admin1"

        'The geography of a linelist is optional, and GeoObject raises when the
        'designer carries none. The suppression covers the lookup alone.
        On Error Resume Next
            Set geo = this.specs.GeoObject()
        On Error GoTo 0

        If geo Is Nothing Then
            LogCheck "Variable " & this.varName & " is a geo variable and the linelist has " & _
                     "no geography, so its dropdowns were skipped", checkingError
            Exit Sub
        End If

        Set cat = geo.GeoLevel(LevelAdmin1)
        drop.Add cat, "admin1"
        cat.Clear
        cat.Push ""

        drop.Add cat, "admin2"
        drop.SetValidation cellRng:=cellRng.Offset(, 1), listName:="admin2", _
            alertType:=varAlert, message:=varMessage

        drop.Add cat, "admin3"
        drop.SetValidation cellRng:=cellRng.Offset(, 2), listName:="admin3", _
            alertType:=varAlert, message:=varMessage

        drop.Add cat, "admin4"
        drop.SetValidation cellRng:=cellRng.Offset(, 3), listName:="admin4", _
            alertType:=varAlert, message:=varMessage
    End Select

    ' Set dropdown validation on the main cell
    drop.SetValidation cellRng:=cellRng, listName:=dropName, _
                       alertType:=varAlert, message:=varMessage, _
                       showError:=(varControl <> "choice_multiple")
End Sub

'@section Data Validation
'===============================================================================

'@label:add-validation
'@sub-title Add min/max data validation to the variable cell
'@details
'Reads min and max values from the dictionary, converts them to Excel
'validation formulas via the ExcelFormula helper, and applies whole-number,
'date, or decimal validation with the configured alert style. Both OS-local
'and standard formula variants are attempted, which works here because
'Validation.Add raises on a range that already carries one, so the first
'success sticks. Text variables are skipped.
Private Sub AddValidation()
    Dim varMin As String
    Dim varMax As String
    Dim varMessage As String
    Dim varAlert As String
    Dim varTypes As String
    Dim validationTypes As Long
    Dim validationAlerts As Long
    Dim excelForm As Variant
    Dim excelOSForm As Variant
    Dim excelForm2 As Variant
    Dim excelOSForm2 As Variant
    Dim cellRng As Range

    varTypes = ValueOf("variable type")
    varMin = ValueOf("min")
    varMax = ValueOf("max")
    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")

    If (varMin = vbNullString And varMax = vbNullString) Or (varTypes = "text") Then Exit Sub

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    With cellRng.Validation
        .Delete

        validationTypes = Switch(varTypes = "integer", xlValidateWholeNumber, _
                                 varTypes = "date", xlValidateDate, _
                                 varTypes = "decimal", xlValidateDecimal)

        validationAlerts = Switch(varAlert = "error", xlValidAlertStop, _
                                  varAlert = "warning", xlValidAlertWarning, _
                                  True, xlValidAlertInformation)

        If (varMin <> vbNullString) And (varMax = vbNullString) Then
            excelForm = ExcelFormula(varMin, validationTypes)
            excelOSForm = ExcelFormula(varMin, validationTypes, useOS:=True)
            If IsEmpty(excelForm) Then Exit Sub

            On Error Resume Next
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlGreaterEqual, Formula1:=excelOSForm
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlGreaterEqual, Formula1:=excelForm
            On Error GoTo 0

        ElseIf (varMin = vbNullString) And (varMax <> vbNullString) Then
            excelForm = ExcelFormula(varMax, validationTypes)
            excelOSForm = ExcelFormula(varMax, validationTypes, useOS:=True)
            If IsEmpty(excelForm) Then Exit Sub

            On Error Resume Next
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlLessEqual, Formula1:=excelOSForm
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlLessEqual, Formula1:=excelForm
            On Error GoTo 0

        ElseIf (varMin <> vbNullString) And (varMax <> vbNullString) Then
            excelForm = ExcelFormula(varMin, validationTypes)
            excelOSForm = ExcelFormula(varMin, validationTypes, useOS:=True)
            excelForm2 = ExcelFormula(varMax, validationTypes)
            excelOSForm2 = ExcelFormula(varMax, validationTypes, useOS:=True)

            If IsEmpty(excelForm) Or IsEmpty(excelForm2) Then Exit Sub

            On Error Resume Next
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlBetween, _
                 Formula1:=excelOSForm, Formula2:=excelOSForm2
            .Add Type:=validationTypes, AlertStyle:=validationAlerts, _
                 Operator:=xlBetween, _
                 Formula1:=excelForm, Formula2:=excelForm2
            On Error GoTo 0
        End If

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .errorTitle = vbNullString
        .InputMessage = vbNullString
        .errorMessage = varMessage
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'@section Conditional Formatting
'===============================================================================

'@label:add-conditional-formatting
'@sub-title Add expression-based conditional formatting
'@details
'Reads the formatting condition variable name from the dictionary, resolves
'its address, and adds an xlExpression format condition that highlights the
'cell when the condition variable equals 1. The interior colour, font
'colour, bold, and italic properties are copied from the formatting values
'cell in the dictionary.
Private Sub AddConditionalFormatting()
    Dim varValid As String
    Dim cellRng As Range
    Dim formatRng As Range
    Dim varObject As LLVariables
    Dim shObj As LLSheets
    Dim condAdd As String
    Dim condForm As String
    Dim OScondForm As String

    varValid = ValueOf("formatting condition")
    Set varObject = Variables()

    If (varValid = vbNullString) Then Exit Sub
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    Set formatRng = varObject.CellRange(colName:="formatting values", _
                                         varName:=this.varName)
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid, this.wksh.Name)

    condForm = "(" & condAdd & "= 1)"
    OScondForm = OSFormula(condForm)

    If Not AddOneCondition(cellRng, OScondForm, "=" & condForm) Then Exit Sub

    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.Color = formatRng.Interior.Color
        .Font.Color = formatRng.Font.Color
        .Font.Bold = formatRng.Font.Bold
        .Font.Italic = formatRng.Font.Italic
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'@label:add-unique-conditional-formatting
'@sub-title Add unique/duplicate conditional formatting (HList only)
'@details
'When the dictionary marks a variable with unique=yes, adds a duplicate-
'values format condition that highlights duplicate entries with a red
'interior. This is typically used for patient IDs or other fields that
'must be unique across records.
Private Sub AddUniqueConditionalFormatting()
    Dim varUnique As String
    Dim cellRng As Range

    varUnique = ValueOf("unique")
    If (varUnique <> "yes") Then Exit Sub

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    cellRng.FormatConditions.AddUniqueValues
    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .DupeUnique = xlDuplicate
        .Interior.Color = vbRed
    End With
End Sub

'@label:add-geo-conditional-formatting
'@sub-title Add geo pcode conditional formatting (HList only)
'@details
'For geo-type variables, adds a format condition that highlights the cell
'in orange when the corresponding pcode variable is empty but the geo
'variable itself is not. This alerts users that a geographic name was
'entered but the pcode lookup failed.
Private Sub AddGeoConditionalFormatting()
    Dim varValid As String
    Dim cellRng As Range
    Dim varObject As LLVariables
    Dim shObj As LLSheets
    Dim condAdd As String
    Dim varControl As String
    Dim actAdd As String
    Dim condForm As String
    Dim OScondForm As String

    varControl = ValueOf("control")
    'Conditional formatting is set only on geo1, geo2, geo3 and geo4 columns. Others
    'should be skept
    If Not (InStr(1, varControl, "geo") = 1) Then Exit Sub

    varValid = "pcode" & "_" & this.varName
    Set varObject = Variables()
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    If cellRng Is Nothing Then Exit Sub

    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid, this.wksh.Name)
    actAdd = shObj.VariableAddress(this.varName, this.wksh.Name)

    condForm = "AND(" & condAdd & "= " & Chr(34) & Chr(34) & "," & actAdd & "<>" & Chr(34) & Chr(34) & ")"
    OScondForm = OSFormula(condForm)

    If Not AddOneCondition(cellRng, OScondForm, "=" & condForm) Then Exit Sub

    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.Color = RGB(242, 147, 12)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'@label:add-one-condition
'@sub-title Give a cell one expression format condition
'@details
'The OS-local formula is tried first and the standard one is used only when
'the first left the count where it was.
'
'`FormatConditions.Add` appends whatever it is given, so the two calls this
'replaces left two conditions on every cell that took both formulas, and only
'the last of them was given a colour, `SetFirstPriority` and `StopIfTrue`. The
'unstyled one then travelled down the ListObject column as rows were added and
'into every exported copy. `Validation.Add` behaves the other way -- it raises
'on a range that already carries one -- which is why `AddValidation` keeps the
'two-call shape.
'@param cellRng Range. The cell to give the condition to.
'@param osFormula String. The OS-local formula, tried first.
'@param standardFormula String. The standard formula, tried only if the first did not take.
'@return Boolean. True when the cell carries a new condition.
Private Function AddOneCondition(ByVal cellRng As Range, _
                                 ByVal osFormula As String, _
                                 ByVal standardFormula As String) As Boolean
    Dim conditionsBefore As Long

    conditionsBefore = cellRng.FormatConditions.Count

    On Error Resume Next
        cellRng.FormatConditions.Add Type:=xlExpression, Formula1:=osFormula
    On Error GoTo 0

    If cellRng.FormatConditions.Count = conditionsBefore Then
        On Error Resume Next
            cellRng.FormatConditions.Add Type:=xlExpression, Formula1:=standardFormula
        On Error GoTo 0
    End If

    AddOneCondition = (cellRng.FormatConditions.Count > conditionsBefore)

    If Not AddOneCondition Then
        LogCheck "Neither conditional formatting formula was accepted for variable " & _
                 this.varName, checkingWarning
    End If
End Function

'@section CRF Companion
'===============================================================================

'@label:add-crf-label
'@sub-title Write CRF label for the current variable
'@details
'Writes the main label and sublabel to the CRF row, applies main-label
'and sublabel character formatting, and hides the row for formula-based
'or hidden variables. Also sets the table anchor named range on the first
'CRF row.
'@param crfRng Range. The target cell on the CRF companion worksheet.
Private Sub AddCRFLabel(ByVal crfRng As Range)
    Dim design As LLFormat
    Dim varMainLab As String
    Dim varSubLab As String
    Dim varControl As String
    Dim varStatus As String
    Dim mainLabLength As Long
    Dim hiddenRow As Boolean
    Dim tabName As String
    Dim crfLine As Long

    Set design = DesignFormat()

    varMainLab = ValueOf("main label")
    varSubLab = ValueOf("sub label")
    varControl = ValueOf("control")
    varStatus = ValueOf("status")
    tabName = ValueOf("table name")
    crfLine = NumberOf("crf index")

    ' Add anchor on first CRF row
    If crfLine = 6 Then crfRng.Name = tabName & "_CRFSTART"

    mainLabLength = Len(varMainLab)
    If (varSubLab <> vbNullString) Then
        varMainLab = varMainLab & Chr(10) & varSubLab
    End If

    ' Write label and apply formatting
    crfRng.Value = varMainLab
    design.ApplyFormat crfRng.Characters(Start:=1, Length:=mainLabLength), scope:=HListMainLab
    design.ApplyFormat crfRng.Characters(Start:=mainLabLength + 1), scope:=HListSublab

    ' Hide rows for formula/hidden variables
    hiddenRow = (varStatus = "hidden" Or varStatus = "optional, hidden") Or _
                (varControl = "formula" Or varControl = "choice_formula" Or _
                 varControl = "case_when")
    crfRng.EntireRow.Hidden = hiddenRow
End Sub

'@label:add-crf-type
'@sub-title Apply CRF-specific type formatting
'@details
'Maps the variable type and format to a CRF design scope constant (such
'as CRFText, CRFDate, CRFDecimal) and applies the corresponding visual
'format to the CRF cell. Skips choice_manual variables because their
'formatting is handled by AddCRFChoices instead.
'@param crfRng Range. The target cell on the CRF companion worksheet.
Private Sub AddCRFType(ByVal crfRng As Range)
    Dim varTypes As String
    Dim varFormat As String
    Dim varControl As String
    Dim scope As Long
    Dim nbDecimals As Integer
    Dim design As LLFormat

    Set design = DesignFormat()

    varTypes = LCase(ValueOf("variable type"))
    varFormat = Trim$(LCase(ValueOf("variable format")))
    varControl = ValueOf("control")
    nbDecimals = 0

    If (varControl = "choice_manual") Then Exit Sub

    Select Case varTypes

    Case "text"
        scope = CRFText

    Case "date"
        scope = IIf(varFormat = vbNullString, CRFDate, CRFFormatCustom)

    Case "decimal", "integer"
        If (varFormat = vbNullString) Then
            scope = IIf(varTypes = "decimal", CRFDecimal, CRFInteger)
        ElseIf (varFormat = "integer") Or (varFormat = "round0") Then
            scope = CRFInteger
        ElseIf (varFormat = "percentage0") Then
            scope = CRFPercentage
        ElseIf (InStr(1, varFormat, "round") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "round", vbNullString))
            On Error GoTo 0
            scope = CRFDecimal
        ElseIf (InStr(1, varFormat, "percentage") > 0) Then
            On Error Resume Next
            nbDecimals = CInt(Replace(varFormat, "percentage", vbNullString))
            On Error GoTo 0
            scope = CRFPercentage
        Else
            scope = CRFFormatCustom
        End If
    End Select

    design.ApplyFormat crfRng, scope:=CByte(scope), nbDecimals:=nbDecimals
End Sub

'@label:add-crf-choices
'@sub-title Write CRF choice headers and value cells
'@details
'For choice_manual variables, retrieves the short-label categories from
'specs, writes optional column headers above the CRF row, and applies
'CRFChoiceValues formatting to each value cell. The headers are only
'written when the "crf choices" dictionary value is "yes".
'@param crfRng Range. The target cell on the CRF companion worksheet.
Private Sub AddCRFChoices(ByVal crfRng As Range)
    Dim choiceHeader As String
    Dim cat As BetterArray
    Dim cellRng As Range
    Dim counter As Long
    Dim varControl As String
    Dim design As LLFormat

    varControl = ValueOf("control")
    If (varControl <> "choice_manual") Then Exit Sub

    Set design = DesignFormat()
    Set cat = this.specs.Categories(this.varName, useShortlabels:=True)
    cat.LowerBound = 1
    If cat.Length = 0 Then Exit Sub

    choiceHeader = ValueOf("crf choices")

    ' Add headers successively
    If (choiceHeader = "yes") Then
        Set cellRng = crfRng.Offset(-2, 1)
        For counter = 1 To cat.Length
            cellRng.Value = cat.Item(counter)
            design.ApplyFormat cellRng, scope:=CRFChoiceHeader
            Set cellRng = cellRng.Offset(, 2)
        Next
    End If

    ' Add choice value cells
    Set cellRng = crfRng.Offset(, 1)
    For counter = 1 To cat.Length
        design.ApplyFormat cellRng, scope:=CRFChoiceValues
        Set cellRng = cellRng.Offset(, 2)
    Next
End Sub

'@section Formula Helpers
'===============================================================================

'@label:scratch-range
'@prop-title The cell a formula is tried on
'@details
'`FailedFormula` and `OSFormula` both need a cell to write a formula into and
'read back. It is A1 of the target worksheet, and both put back what they found.
'
'WHY A1 AND NOT A CELL OUT OF THE WAY
'-------------------------------------------------------------------------------
'A far cell looks safer and it is not: writing to the last column stretches the
'worksheet's used range across all 16,384 columns, and from then on every
'`EntireColumn`, every merge and every full-column read on that sheet works over
'that range. A sheet build that took a minute did not finish in nine. The cost
'is Excel's bookkeeping, so it does not show up in the code that pays it.
'
'What was actually wrong with A1 was `FailedFormula` reaching it with `.Clear`,
'which takes the formatting as well as the contents. Both helpers save the
'formula and put it back now.
'@return Range. The scratch cell.
Private Property Get ScratchRange() As Range
    If this.scratchCell Is Nothing Then
        Set this.scratchCell = this.wksh.Cells(1, 1)
    End If
    Set ScratchRange = this.scratchCell
End Property

'@label:failed-formula
'@sub-title Test if a formula fails when applied to a cell
'@details
'Writes the formula to the scratch cell as a test. If the assignment raises an
'error the function returns True, indicating the formula is invalid. The cell
'contents are cleared after the test whatever the outcome, and
'Application.DisplayAlerts is suppressed for the duration.
'@param frmName String. The Excel formula expression to test (without leading equals sign).
'@return Boolean. True if the formula raised an error, False otherwise.
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim scratchRng As Range
    Dim prevAlert As Boolean
    Dim prevForm As String
    Dim didFail As Boolean

    Set scratchRng = ScratchRange()
    prevAlert = Application.DisplayAlerts
    prevForm = scratchRng.formula

    Application.DisplayAlerts = False

    On Error GoTo TestFailed
    scratchRng.formula = "= " & frmName
    On Error GoTo 0

CleanUpTest:
    'The cleanup sits under its own suppression and is reached by Resume, so a
    'second failure here cannot end the call the way an inline restore inside an
    'armed handler would.
    On Error Resume Next
        scratchRng.formula = prevForm
        Application.DisplayAlerts = prevAlert
    On Error GoTo 0

    FailedFormula = didFail
    Exit Function

TestFailed:
    didFail = True
    Resume CleanUpTest
End Function

'@label:excel-formula
'@sub-title Convert a formula value to an Excel-compatible format for validation
'@details
'Parses the given formula string through the Formulas class, optionally
'converts it to an OS-local formula via OSFormula, and then converts the
'result to the type the validation asks for. A whole number takes `CLng` and a
'decimal takes `CDbl`: `CInt` overflowed above 32,767, which is a legitimate
'bound for a count or an ID, and `CLng` on a decimal turned a maximum of 0.5
'into 0, so the validation rejected every value the dictionary meant to allow.
'Returns Empty when the formula is invalid, fails the FailedFormula test, or
'will not convert -- and the last of those is filed through Checking, because
'the caller runs under a suppression and would otherwise drop the validation
'with no message.
'@param formVal String. The raw formula value from the dictionary (e.g. a min or max expression).
'@param validationTypes Long. The Excel validation type constant (xlValidateWholeNumber, xlValidateDate, or xlValidateDecimal).
'@param useOS Optional Boolean. When True, converts the formula to OS-local format via OSFormula. Defaults to False.
'@return Variant. The converted formula value, or Empty if invalid.
Private Function ExcelFormula(ByVal formVal As String, _
                              ByVal validationTypes As Long, _
                              Optional ByVal useOS As Boolean = False) As Variant
    Dim dict As LLdictionary
    Dim varFormData As FormulaData
    Dim varFormObject As Formulas
    Dim excelForm As String
    Dim convForm As Variant

    Set dict = Dictionary()
    Set varFormData = FormulaDataObj()
    Set varFormObject = Formulas.Create(dict, varFormData, formVal)

    If Not varFormObject.Valid(formulaType:="linelist") Then
        If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
        LogCheck "The validation bound " & formVal & " of variable " & this.varName & _
                 " was rejected by the parser, so the validation was left off", checkingError
        Exit Function
    End If

    excelForm = Trim$(Replace(varFormObject.ParsedLinelistFormula(), "=", vbNullString))

    'A bound that reads as a number goes to the validation as a number. Excel is
    'asked to judge nothing, which matters because the parser reads a decimal in
    'the host's separator and Range.Formula reads one in US syntax -- so probing
    'a bound of 0,5 by writing it into a cell declared every decimal bound
    'broken, and AddValidation runs under a suppression, so the variable lost its
    'validation with no message.
    If Not varFormObject.HasSetupVariables Then
        convForm = ConvertedBound(excelForm, validationTypes)
        If Not IsEmpty(convForm) Then
            If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
            ExcelFormula = convForm
            Exit Function
        End If
    End If

    'Anything else is handed over as formula text, and Excel is the judge of it.
    If FailedFormula(excelForm) Then
        If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
        LogCheck "The validation bound " & formVal & " of variable " & this.varName & _
                 " was refused by Excel, so the validation was left off", checkingError
        Exit Function
    End If

    If useOS Then
        excelForm = OSFormula(excelForm)
    Else
        excelForm = "= " & excelForm
    End If

    ExcelFormula = excelForm

    ' Harvest formula parsing checkings
    If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
End Function

'@label:converted-bound
'@sub-title Read a bound as the type the validation asks for
'@details
'Answers Empty when the text will not convert, which is how the caller knows to
'hand the bound over as formula text instead. A whole number takes `CLng` and a
'decimal takes the separator-independent read: `CInt` overflowed above 32,767,
'which is a legitimate bound for a count or an ID, and `CLng` on a decimal
'turned a maximum of 0.5 into 0, so the validation rejected every value the
'dictionary meant to allow.
'@param boundText String. The parsed bound.
'@param validationTypes Long. The Excel validation type constant.
'@return Variant. The converted bound, or Empty.
Private Function ConvertedBound(ByVal boundText As String, _
                                ByVal validationTypes As Long) As Variant
    On Error GoTo NotANumber

    Select Case validationTypes
    Case xlValidateWholeNumber
        ConvertedBound = CLng(AsNumber(boundText))
    Case xlValidateDate
        ConvertedBound = CDate(boundText)
    Case xlValidateDecimal
        ConvertedBound = AsNumber(boundText)
    End Select

    Exit Function

NotANumber:
End Function

'@label:as-number
'@sub-title Read a bound as a number whatever decimal separator it carries
'@details
'`CDbl` reads the separator of the host, and a dictionary is authored on
'whatever machine the user has. A bound typed as `0.5` on one and read back on
'a comma machine is text `CDbl` refuses, and a bound of `0,5` is text a dot
'machine refuses. The separator is taken from the host and both are mapped onto
'it.
'
'`Val` was the other candidate and it answers 0 for `0,5`, which is the same
'silent truncation this method exists to stop.
'@param numberText String. The bound as the dictionary carries it.
'@return Double. The number.
Private Function AsNumber(ByVal numberText As String) As Double
    Dim hostSeparator As String
    Dim normalised As String

    hostSeparator = Mid$(CStr(1.5), 2, 1)

    normalised = Replace(numberText, ".", hostSeparator)
    normalised = Replace(normalised, ",", hostSeparator)

    AsNumber = CDbl(normalised)
End Function

'@label:os-formula
'@sub-title Convert a formula to OS-specific local format
'@details
'On macOS with VList layers, returns the formula as-is because FormulaLocal
'is unreliable. On Windows or for HList layers, temporarily writes the
'formula to the scratch cell, reads back FormulaLocal, and restores the
'previous cell content. This ensures validation formulas use the correct
'locale-aware separator characters.
'
'The Mac carve-out covers VList alone and the HList branch calls FormulaLocal
'whatever the machine is. If FormulaLocal is unreliable on Mac for one layer it
'is unreliable for the other, and that is still open.
'@param formVal String. The Excel formula expression to convert.
'@return String. The OS-local formula string.
Private Function OSFormula(ByVal formVal As String) As String
    Dim scratchRng As Range
    Dim prevForm As String

    OSFormula = vbNullString

    If (this.layer = VarWriterLayerVList) And (Application.OperatingSystem Like "*Mac*") Then
        OSFormula = formVal
        Exit Function
    End If

    Set scratchRng = ScratchRange()
    prevForm = scratchRng.formula
    scratchRng.formula = "=" & formVal
    OSFormula = scratchRng.FormulaLocal
    scratchRng.formula = prevForm
End Function

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise an error with the class name as source
'@details
'Raises the number it is given with the class name as the source. Every guard
'in this class names a `ProjectError` member.
'@param errNumber Long. The ProjectError value to raise.
'@param errorMessage String. The error description to include in the raised error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, "Error : " & errorMessage
End Sub


'@section Checkings
'===============================================================================
'@description
'What this class files is a problem: a variable it could not place, a formula
'the parser rejected, a validation bound that would not convert, a dropdown
'manager it was not given. The per-variable "written" entry that used to be
'filed here is gone -- a 500-variable linelist produced 500 report rows saying
'nothing, and each of them cloned the whole key array.
'
'KEYS ARE UNIQUE ACROSS A BUILD
'-------------------------------------------------------------------------------
'`Checking.Add` raises on a duplicate key. Every entry filed here is keyed by
'the worksheet and the variable, so a second subsection, a second sheet and a
'second writer can all be merged into one Checking. The counter this replaces
'started at 0 in every instance, so the second subsection of a sheet re-added
'key "0" and stopped the build.
'
'Collaborator entries are re-keyed on the way in for the same reason:
'`Formulas` also keys with a counter from 0, and `AddFormula` builds a fresh
'`Formulas` per variable.

'@label:log-check
'@sub-title File one entry against the current variable
'@param label String. The message for the report.
'@param scope Optional Byte. Severity. Defaults to checkingWarning.
Private Sub LogCheck(ByVal label As String, Optional ByVal scope As Byte = checkingWarning)
    EnsureChecking
    internalChecks.Add UniqueKey(EntryKeyRoot()), label, scope
    this.hasChecks = True
End Sub

'@label:harvest-collaborator-checkings
'@sub-title Merge another object's entries into this writer's store
'@details
'The entries are re-keyed under this variable rather than appended, so two
'collaborators that both key from 0 do not collide. The severity is read back
'through the type label `Checking` prints, which is the only way it is exposed.
'@param collaboratorChecks Checking. The source store.
Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As Checking)
    Dim sourceKeys As BetterArray
    Dim counter As Long
    Dim sourceKey As String

    If collaboratorChecks Is Nothing Then Exit Sub
    If collaboratorChecks.Length() <= 0 Then Exit Sub

    EnsureChecking
    Set sourceKeys = collaboratorChecks.ListOfKeys()
    If sourceKeys Is Nothing Then Exit Sub

    For counter = sourceKeys.LowerBound To sourceKeys.UpperBound
        sourceKey = CStr(sourceKeys.Item(counter))
        internalChecks.Add UniqueKey(EntryKeyRoot()), _
                           collaboratorChecks.ValueOf(sourceKey, checkingLabel), _
                           ScopeOf(collaboratorChecks, sourceKey)
    Next counter

    this.hasChecks = True
End Sub

'@label:ensure-checking
'@sub-title Build the internal store on first use
Private Sub EnsureChecking()
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("Variables")
    End If
End Sub

'@label:entry-key-root
'@sub-title The key an entry of the current variable starts from
'@return String. Worksheet name, then the variable name.
Private Function EntryKeyRoot() As String
    EntryKeyRoot = this.wksh.Name & "!" & this.varName
End Function

'@label:unique-key
'@sub-title Answer a key the internal store does not hold yet
'@details
'A variable can file more than one entry -- a rejected formula and a validation
'bound that would not convert, for instance -- so the root is numbered when it
'is already taken.
'@param keyRoot String. The key to start from.
'@return String. A key that is free.
Private Function UniqueKey(ByVal keyRoot As String) As String
    Dim candidate As String
    Dim suffix As Long

    candidate = keyRoot
    Do While internalChecks.KeyExists(candidate)
        suffix = suffix + 1
        candidate = keyRoot & " #" & CStr(suffix)
    Loop

    UniqueKey = candidate
End Function

'@label:scope-of
'@sub-title Read the severity of one entry of another checking
'@details
'`Checking` exposes a scope as the decorated type label it prints in the
'report, so the byte is recovered from that text. Each of the five labels
'carries its own word after the icon, and matching on the word keeps this free
'of the icon code points.
'@param source Checking. The store holding the entry.
'@param keyName String. The entry key.
'@return Byte. A CheckingScope member.
Private Function ScopeOf(ByVal source As Checking, ByVal keyName As String) As Byte
    Dim typeLabel As String

    typeLabel = source.ValueOf(keyName, checkingType)

    If InStr(1, typeLabel, "Error", vbTextCompare) > 0 Then
        ScopeOf = checkingError
    ElseIf InStr(1, typeLabel, "Warning", vbTextCompare) > 0 Then
        ScopeOf = checkingWarning
    ElseIf InStr(1, typeLabel, "Info", vbTextCompare) > 0 Then
        ScopeOf = checkingInfo
    ElseIf InStr(1, typeLabel, "Success", vbTextCompare) > 0 Then
        ScopeOf = checkingSuccess
    Else
        ScopeOf = checkingNote
    End If
End Function

Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property
