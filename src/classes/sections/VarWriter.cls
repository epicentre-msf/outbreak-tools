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
Implements IVarWriter

'@class VarWriter
'@description
'Writes a single dictionary variable to one or more linelist worksheets.
'For HList layers the writer targets the main sheet, a printed companion,
'and a CRF companion; for VList layers only the main sheet is written.
'Each invocation of WriteVariable applies label text, number formatting,
'formulas, dropdown choices, data validation, and conditional formatting
'to the resolved cell range. The class caches dictionary, variables, and
'sheets accessors for performance across repeated calls.
'@depends ILinelistSpecs, ILLdictionary, ILLVariables, ILLSheets, ILLFormat,
'ITranslationObject, IFormulaData, IFormulas, IDropdownLists, ILLGeo,
'Formulas, LLVariables, LLSheets, BetterArray
'@version 1.0 (2026-02-10)


Private Type TVarWriter
    layer As Byte
    specs As ILinelistSpecs
    wksh As Worksheet
    hn As IHiddenNames
    printWksh As Worksheet
    crfWksh As Worksheet
    dropdown As IDropdownLists
    customDropdown As IDropdownLists
    varName As String
    cachedRange As Range
    cachedPrintRange As Range
    cachedCRFRange As Range
    cachedDict As ILLdictionary
    cachedVars As ILLVariables
    cachedSheets As ILLSheets
    hasChecks As Boolean
End Type

Private this As TVarWriter
Private Const CLASS_NAME As String = "VarWriter"
Private Const HLIST_START_ROW As Long = 8
Private Const VLIST_START_COL As Long = 5

'Checking support
Private checkCounter As Long
Private internalChecks As IChecking

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new VarWriter for the given layer and dependencies
'@details
'Factory method that instantiates and configures a VarWriter. The caller
'must supply a valid ILinelistSpecs and target Worksheet at minimum.
'Optional parameters allow attaching printed and CRF companion worksheets
'(used only for HList layers) and dropdown list managers. Raises an error
'if specs or wksh is Nothing.
'@param layer Byte. VarWriterLayer enum value (HList=1, VList=2). Use Byte for macOS compatibility.
'@param specs ILinelistSpecs. The linelist specifications object providing dictionary, format, translations, etc.
'@param wksh Worksheet. The target main worksheet.
'@param printWksh Optional Worksheet. Printed companion worksheet (HList only). Defaults to Nothing.
'@param crfWksh Optional Worksheet. CRF companion worksheet (HList only). Defaults to Nothing.
'@param dropdownObj Optional IDropdownLists. Dropdown lists manager. Defaults to Nothing.
'@param customDropdownObj Optional IDropdownLists. Custom dropdown lists manager (for choice_custom). Defaults to Nothing.
'@return IVarWriter. A fully initialised writer bound to the given layer and worksheets.
'@throws VBA.vbObjectError + 516 if specs or wksh is Nothing.
'@export
Public Function Create(ByVal layer As Byte, _
                       ByVal specs As ILinelistSpecs, _
                       ByVal wksh As Worksheet, _
                       Optional ByVal printWksh As Worksheet = Nothing, _
                       Optional ByVal crfWksh As Worksheet = Nothing, _
                       Optional ByVal dropdownObj As IDropdownLists = Nothing, _
                       Optional ByVal customDropdownObj As IDropdownLists = Nothing) As IVarWriter

    If specs Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Specifications object is required"
    If wksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Target worksheet is required"

    With New VarWriter
        Set .InternalSpecs = specs
        Set .InternalWksh = wksh
        .InternalLayer = layer
        If Not printWksh Is Nothing Then Set .InternalPrintWksh = printWksh
        If Not crfWksh Is Nothing Then Set .InternalCRFWksh = crfWksh
        If Not dropdownObj Is Nothing Then Set .InternalDropdown = dropdownObj
        If Not customDropdownObj Is Nothing Then Set .InternalCustomDropdown = customDropdownObj
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Return the current instance cast to the IVarWriter interface
'@details
'Used by the factory pattern so that the With New block can return an
'interface-typed reference instead of the concrete class.
'@return IVarWriter. The current instance.
Public Property Get Self() As IVarWriter
    Set Self = Me
End Property

'@section Internal Setters
'@description Properties used by the factory only during instantiation.
'===============================================================================

'@label:internal-specs
'@prop-title Assign the linelist specifications object
Public Property Set InternalSpecs(ByVal specs As ILinelistSpecs)
    Set this.specs = specs
End Property

'@label:internal-wksh
'@prop-title Assign the target main worksheet
Public Property Set InternalWksh(ByVal wksh As Worksheet)
    Set this.wksh = wksh
    Set this.hn = HiddenNames.Create(this.wksh)
End Property

'@label:internal-layer
'@prop-title Assign the writer layer (HList or VList)
Public Property Let InternalLayer(ByVal layer As Byte)
    this.layer = layer
End Property

'@label:internal-print-wksh
'@prop-title Assign the printed companion worksheet
Public Property Set InternalPrintWksh(ByVal wksh As Worksheet)
    Set this.printWksh = wksh
End Property

'@label:internal-crf-wksh
'@prop-title Assign the CRF companion worksheet
Public Property Set InternalCRFWksh(ByVal wksh As Worksheet)
    Set this.crfWksh = wksh
End Property

'@label:internal-dropdown
'@prop-title Assign the dropdown lists manager
Public Property Set InternalDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.dropdown = dropdownObj
End Property

'@label:internal-custom-dropdown
'@prop-title Assign the custom dropdown lists manager
Public Property Set InternalCustomDropdown(ByVal dropdownObj As IDropdownLists)
    Set this.customDropdown = dropdownObj
End Property

'@section Cached Accessors
'@description Lazy-loaded accessors for frequently used dependency objects.
'===============================================================================

'@label:dictionary
'@prop-title Return the cached dictionary object
'@details
'Lazily loads the ILLdictionary from specs on first access and caches it
'in the TVarWriter UDT. Subsequent calls return the cached instance
'without re-querying specs, which improves performance when writing many
'variables in sequence.
'@return ILLdictionary. The linelist dictionary.
Private Property Get Dictionary() As ILLdictionary
    If this.cachedDict Is Nothing Then
        Set this.cachedDict = this.specs.Dictionary()
    End If
    Set Dictionary = this.cachedDict
End Property

'@label:variables
'@prop-title Return the cached variables helper
'@details
'Creates an LLVariables wrapper around the dictionary on first access.
'The wrapper provides convenient lookup methods such as Value and Contains
'that are used throughout the writer.
'@return ILLVariables. The variables helper object.
Private Property Get Variables() As ILLVariables
    If this.cachedVars Is Nothing Then
        Set this.cachedVars = LLVariables.Create(Dictionary())
    End If
    Set Variables = this.cachedVars
End Property

'@label:sheets-object
'@prop-title Return the cached sheets helper
'@details
'Creates an LLSheets wrapper around the dictionary on first access.
'The wrapper provides VariableAddress and SheetInfo lookups used when
'building conditional-formatting formulas and resolving table anchors.
'@return ILLSheets. The sheets helper object.
Private Property Get SheetsObject() As ILLSheets
    If this.cachedSheets Is Nothing Then
        Set this.cachedSheets = LLSheets.Create(Dictionary())
    End If
    Set SheetsObject = this.cachedSheets
End Property

'@label:design-format
'@prop-title Return the design format object from specs
'@details
'Delegates directly to specs.DesignFormat on every call because the
'format object is lightweight and does not benefit from caching.
'@return ILLFormat. The design format provider.
Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = this.specs.DesignFormat()
End Property

'@label:trans
'@prop-title Return the translation object from specs
'@details
'Delegates directly to specs.TransObject on every call. Used to fetch
'translated UI messages such as MSG_Calculated and MSG_CustomChoice.
'@return ITranslationObject. The translation provider.
Private Property Get Trans() As ITranslationObject
    Set Trans = this.specs.TransObject()
End Property

Private Property Get HNames() As IHiddenNames
    Set HNames = this.hn
End Property

'@label:formula-data-obj
'@prop-title Return the formula data object from specs
'@details
'Delegates directly to specs.FormulaDataObject on every call. Provides
'the lookup tables and metadata needed by the Formulas parser.
'@return IFormulaData. The formula data provider.
Private Property Get FormulaDataObj() As IFormulaData
    Set FormulaDataObj = this.specs.FormulaDataObject()
End Property

'@section Value Lookup
'===============================================================================

'@label:value-of
'@prop-title Read a dictionary column value for the current variable
'@details
'Looks up the value stored in the specified dictionary column for the
'variable that was most recently passed to WriteVariable. This is the
'primary way all other methods retrieve variable metadata such as labels,
'types, formats, and control information.
'@param colName String. The dictionary column name to look up.
'@return String. The column value for the current variable.
'@export
Public Property Get ValueOf(ByVal colName As String) As String
    ValueOf = Variables().Value(colName:=colName, varName:=this.varName)
End Property

'@section Cell Range Resolution
'===============================================================================

'@label:var-range
'@prop-title Resolve the main cell range for the current variable
'@details
'Computes the target cell from the column index dictionary value and the
'layer-specific start position. For HList the cell is in a column at
'HLIST_START_ROW + 1; for VList it is in a row at VLIST_START_COL. The
'result is cached until WriteVariable resets it for the next variable.
'@return Range. The main cell range, or Nothing if the index is invalid.
Private Property Get VarRange() As Range
    If this.cachedRange Is Nothing Then
        Dim idx As Long
        idx = CLng(ValueOf("column index"))

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
            idx = CLng(ValueOf("column index"))
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
            On Error Resume Next
            crfLine = CLng(ValueOf("crf index"))
            On Error GoTo 0

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
'resets the cached cell ranges, then dispatches to WriteMainSheet. For
'HList layers it additionally dispatches to WritePrintSheet and
'WriteCRFSheet for the companion worksheets.
'@param varName String. The variable name as it appears in the dictionary.
'@export
Public Sub WriteVariable(ByVal varName As String)
    this.varName = varName

    ' Reset cached ranges for new variable
    Set this.cachedRange = Nothing
    Set this.cachedPrintRange = Nothing
    Set this.cachedCRFRange = Nothing

    ' Write to main sheet
    WriteMainSheet

    ' HList: also write to printed and CRF companions
    If this.layer = VarWriterLayerHList Then
        WritePrintSheet
        WriteCRFSheet
    End If

    LogInfo "Variable " & varName & " written"
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
    Dim calculated As String
    Dim mandatory As String
    Dim mainLabLength As Long
    Dim design As ILLFormat
    Dim varToPrint As String
    Dim choicust As String
    Dim varEditLab As String
    Dim transObj As ITranslationObject

    Set design = DesignFormat()
    Set transObj = Trans()

    varMainLab = ValueOf("main label")
    varName = ValueOf("variable name")
    varNote = ValueOf("note")
    varSubLab = ValueOf("sub label")
    varControl = ValueOf("control")
    varEditLab = ValueOf("editable label")
    varStatus = ValueOf("status")
    varListAutoOrigin = ValueOf("list auto")
    varIndex = CLng(ValueOf("column index"))
    calculated = transObj.TranslatedValue("MSG_Calculated")
    mandatory = transObj.TranslatedValue("MSG_Mandatory")
    choicust = transObj.TranslatedValue("MSG_CustomChoice")

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
        varSubLab = varSubLab & Chr(10) & calculated
    End If

    If (varControl = "choice_custom") Then
        varSubLab = varSubLab & Chr(10) & "(" & choicust & ")"
    End If

    mainLabLength = Len(varMainLab)
    If (varSubLab <> vbNullString) Then
        varMainLab = varMainLab & Chr(10) & varSubLab
    End If

    ' Status to add to notes
    If varStatus = "mandatory" Then varNote = mandatory & Chr(10) & varNote

    ' --- HList-specific label writing ---
    If this.layer = VarWriterLayerHList Then
        ' Table start anchor
        If (varIndex = 1) And (Not usePrinted) Then
            cellRng.Offset(-1).Name = ValueOf("table name") & "_START"
        End If
        If (varIndex = 1) And usePrinted Then
            cellRng.Offset(-1).Name = ValueOf("table name") & "_PRINTSTART"
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

        ' Variable notes
        If varNote <> vbNullString And (Not usePrinted) Then
            cellRng.Offset(-2).AddComment
            cellRng.Offset(-2).Comment.Text Text:=varNote
            cellRng.Offset(-2).Comment.Visible = False
        End If

        ' Control info in hidden rows - Use HiddenNames instead
        HNames.EnsureName varName & " -- " & "control", varControl, HiddenNameTypeString
        HNames.EnsureName varName & " -- " & "listauto", varListAutoOrigin, HiddenNameTypeString
        HNames.EnsureName varName & " -- " & "address", cellRng.Address, HiddenNameTypeString

        ' Sheet-level flag for listauto presence (used by EventLinelist)
        If varListAutoOrigin <> vbNullString Then
            HNames.EnsureName "has_listauto", "yes", HiddenNameTypeString
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
            HNames.EnsureName varName & " -- " & "editable", "yes", HiddenNameTypeString
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
        ' Table start anchor
        If varIndex = 4 Then cellRng.Name = ValueOf("table name") & "_START"

        ' Main label and sub label
        cellRng.Offset(, -1).Value = varMainLab

        If varControl <> "choice_custom" Then
            design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=1, Length:=mainLabLength), _
                scope:=VListMainLab
            design.ApplyFormat cellRng.Offset(, -1).Characters(Start:=mainLabLength + 1), _
                scope:=VListSublab
        End If

        ' Variable name
        HNames.EnsureName varName & " -- " & "address", cellRng.Address, HiddenNameTypeString
        cellRng.Name = varName

        ' Variable note
        If varNote <> vbNullString Then
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
            HNames.EnsureName varName & " -- " & "editable", "yes", HiddenNameTypeString
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
'euros, and dollars as well as custom format strings.
Private Sub AddType()
    Dim varTypes As String
    Dim varFormat As String
    Dim cellRng As Range
    Dim nbDecimals As Integer
    Dim zerosDecimals As String
    Dim formatText As String

    Set cellRng = VarRange()
    varTypes = LCase(Application.WorksheetFunction.Trim(ValueOf("variable type")))
    varFormat = LCase(Application.WorksheetFunction.Trim(ValueOf("variable format")))
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
            formatText = "#,##0.00 [$" & Chr(8364) & "-x-euro1];-#,##0.00 [$" & Chr(8364) & "-x-euro1];0.00 [$" & Chr(8364) & "-x-euro1];@"
        ElseIf varFormat = "dollars" Then
            formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"
        ElseIf varFormat <> "text" Then
            formatText = "custom"
        End If

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
'and writes the Excel formula to the cell. If the parsed formula fails a
'test application it is replaced with a failure placeholder string. Also
'applies calculated-formula design formatting to the cell and its header.
Private Sub AddFormula()
    Dim varForm As String
    Dim excelForm As String
    Dim cellRng As Range
    Dim varFormData As IFormulaData
    Dim dict As ILLdictionary
    Dim varFormObject As IFormulas
    Dim design As ILLFormat
    Dim varControl As String

    varControl = ValueOf("control")
    If (varControl <> "formula") And (varControl <> "case_when") And _
       (varControl <> "choice_formula") Then Exit Sub

    Set dict = Dictionary()
    Set varFormData = FormulaDataObj()
    Set design = DesignFormat()
    varForm = ValueOf("control details")
    Set cellRng = VarRange()

    Set varFormObject = Formulas.Create(dict, varFormData, varForm)

    If varFormObject.Valid(formulaType:="linelist") Then
        excelForm = varFormObject.ParsedLinelistFormula()
        If FailedFormula(excelForm) Then
            excelForm = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
        End If
        cellRng.formula = "= " & excelForm
        cellRng.Locked = True
    End If

    ' Apply calculated formula format
    design.ApplyFormat cellRng, scope:=HListCalculatedFormulaCell
    If this.layer = VarWriterLayerHList Then
        design.ApplyFormat cellRng.Offset(-2), scope:=HListCalculatedFormulaHeader
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
'registers them with the appropriate IDropdownLists manager, and calls
'SetValidation to apply dropdown data validation to the cell. The
'choice_custom branch additionally creates a hyperlink from the label cell
'to the dropdown worksheet and adds a return link. The geo1 branch sets up
'admin1 through admin4 dropdown cascades.
Private Sub AddChoices()
    Dim varControl As String
    Dim dropName As String
    Dim varAlert As String
    Dim varMessage As String
    Dim cat As BetterArray
    Dim drop As IDropdownLists
    Dim cellRng As Range
    Dim geo As ILLGeo
    Dim choicust As String
    Dim transObj As ITranslationObject
    Dim varMainLab As String
    Dim varSubLab As String
    Dim sh As Worksheet
    Dim labRng As Range
    Dim targetAddress As String
    Dim mainLabLength As Long
    Dim design As ILLFormat


    varControl = ValueOf("control")

    On Error Resume Next
    varControl = Split(varControl, "(")(0)
    On Error GoTo 0

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
    Set sh = cellRng.Parent

    varAlert = ValueOf("alert")
    varMessage = ValueOf("message")
    Set cat = New BetterArray

    Select Case varControl

    Case "choice_manual", "choice_multiple"
        Set drop = this.dropdown
        dropName = this.varName
        Set cat = this.specs.Categories(dropName)
        drop.Add cat, dropName

    Case "choice_custom"
        Set drop = this.customDropdown
        Set transObj = Trans()
        Set design = DesignFormat()
        choicust = transObj.TranslatedValue("MSG_CustomChoice")
        varSubLab = ValueOf("sub label")
        varMainLab = ValueOf("main label")
        mainLabLength = Len(varMainLab)

        dropName = this.varName
        Set cat = this.specs.Categories(dropName)
        If cat.Length = 0 Then cat.Push vbNullString, vbNullString

        drop.Add cat, dropName, addLabel:=True, counterPrefix:=choicust

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
        Set drop = this.dropdown
        dropName = ValueOf("control details")
        cat.Push ""
        drop.Add cat, dropName

    Case "geo1"
        ' HList only
        Set drop = this.dropdown
        dropName = "admin1"
        Set geo = this.specs.GeoObject()
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
'and standard formula variants are attempted to ensure cross-platform
'compatibility. Text variables are skipped.
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
    Dim varObject As ILLVariables
    Dim shObj As ILLSheets
    Dim condAdd As String
    Dim condForm As String
    Dim OScondForm As String

    varValid = ValueOf("formatting condition")
    Set varObject = Variables()

    If (varValid = vbNullString) Then Exit Sub
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    Set formatRng = varObject.CellRange(colName:="formatting values", _
                                         varName:=this.varName)
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid, this.wksh.Name)

    condForm = "(" & condAdd & "= 1)"
    OScondForm = OSFormula(condForm)

    On Error Resume Next
        cellRng.FormatConditions.Add Type:=xlExpression, Formula1:=OScondForm
        cellRng.FormatConditions.Add Type:=xlExpression, Formula1:="=" & condForm
    On Error GoTo 0

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
    Dim varObject As ILLVariables
    Dim shObj As ILLSheets
    Dim condAdd As String
    Dim varControl As String
    Dim actAdd As String
    Dim condForm As String
    Dim OScondForm As String

    varControl = ValueOf("control")
    'Conditional formatting is set only on geo1, geo2, geo3 and geo4 columns. Others
    'should be skept
    If Not (InStr(1, varControl, "geo") = 1)  Then Exit Sub

    varValid = "pcode" & "_" & this.varName
    Set varObject = Variables()
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid, this.wksh.Name)
    actAdd = shObj.VariableAddress(this.varName, this.wksh.Name)
    
    condForm = "AND(" & condAdd & "= " & Chr(34) & Chr(34) & "," & actAdd & "<>" & Chr(34) & Chr(34) & ")"
    OScondForm = OSFormula(condForm)

    On Error Resume Next
        cellRng.FormatConditions.Add Type:=xlExpression, Formula1:=OScondForm
        cellRng.FormatConditions.Add Type:=xlExpression, Formula1:= "=" & condForm
    On Error GoTo 0

    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.Color = RGB(242, 147, 12)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

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
    Dim design As ILLFormat
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

    On Error Resume Next
    crfLine = CLng(ValueOf("crf index"))
    On Error GoTo 0

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
    Dim design As ILLFormat

    Set design = DesignFormat()

    varTypes = LCase(ValueOf("variable type"))
    varFormat = LCase(ValueOf("variable format"))
    varFormat = Application.WorksheetFunction.Trim(varFormat)
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
    Dim design As ILLFormat

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

'@label:failed-formula
'@sub-title Test if a formula fails when applied to a cell
'@details
'Writes the formula to cell A1 of the main worksheet as a test. If the
'assignment raises an error the function returns True, indicating the
'formula is invalid. The cell is cleared after the test regardless of
'outcome. Application.DisplayAlerts is temporarily suppressed during
'the test.
'@param frmName String. The Excel formula expression to test (without leading equals sign).
'@return Boolean. True if the formula raised an error, False otherwise.
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim sh As Worksheet
    Dim prevAlert As Boolean

    Set sh = this.wksh
    With sh
        .Cells(1, 1).Clear
        On Error GoTo Fail
        prevAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False
        .Cells(1, 1).formula = "= " & frmName
        .Cells(1, 1).Clear
        FailedFormula = False
        Application.DisplayAlerts = prevAlert
        Exit Function
    End With

Fail:
    FailedFormula = True
    Application.DisplayAlerts = prevAlert
End Function

'@label:excel-formula
'@sub-title Convert a formula value to an Excel-compatible format for validation
'@details
'Parses the given formula string through the Formulas class, optionally
'converts it to an OS-local formula via OSFormula, and then converts the
'result to the appropriate VBA type (CInt, CDate, or CLng) based on the
'validation type when no formula literals are present. Returns Empty if
'the formula is invalid or fails the FailedFormula test.
'@param formVal String. The raw formula value from the dictionary (e.g. a min or max expression).
'@param validationTypes Long. The Excel validation type constant (xlValidateWholeNumber, xlValidateDate, or xlValidateDecimal).
'@param useOS Optional Boolean. When True, converts the formula to OS-local format via OSFormula. Defaults to False.
'@return Variant. The converted formula value, or Empty if invalid.
Private Function ExcelFormula(ByVal formVal As String, _
                              ByVal validationTypes As Long, _
                              Optional ByVal useOS As Boolean = False) As Variant
    Dim dict As ILLdictionary
    Dim varFormData As IFormulaData
    Dim varFormObject As IFormulas
    Dim excelForm As String
    Dim convForm As Variant

    Set dict = Dictionary()
    Set varFormData = FormulaDataObj()
    Set varFormObject = Formulas.Create(dict, varFormData, formVal)

    If Not varFormObject.Valid(formulaType:="linelist") Then
        If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
        Exit Function
    End If
    excelForm = varFormObject.ParsedLinelistFormula()
    If FailedFormula(excelForm) Then
        If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
        Exit Function
    End If

    If useOS Then
        excelForm = OSFormula(excelForm)
    Else
        excelForm = "= " & excelForm
    End If

    If Not varFormObject.HasSetupVariables Then
        excelForm = Application.WorksheetFunction.Trim(Replace(excelForm, "=", ""))

        Select Case validationTypes
        Case xlValidateWholeNumber
            convForm = CInt(excelForm)
        Case xlValidateDate
            convForm = CDate(excelForm)
        Case xlValidateDecimal
            convForm = CLng(excelForm)
        End Select
        ExcelFormula = convForm
    Else
        ExcelFormula = excelForm
    End If

    ' Harvest formula parsing checkings
    If varFormObject.HasChecking Then HarvestCollaboratorCheckings varFormObject.CheckingValues
End Function

'@label:os-formula
'@sub-title Convert a formula to OS-specific local format
'@details
'On macOS with VList layers, returns the formula as-is because FormulaLocal
'is unreliable. On Windows or for HList layers, temporarily writes the
'formula to cell A1, reads back FormulaLocal, and restores the previous
'cell content. This ensures validation formulas use the correct locale-
'aware separator characters.
'@param formVal String. The Excel formula expression to convert.
'@return String. The OS-local formula string.
Private Function OSFormula(ByVal formVal As String) As String
    Dim sh As Worksheet
    Dim prevForm As String

    OSFormula = vbNullString

    If this.layer = VarWriterLayerVList Then
        ' VList: on Mac, return formula as-is; on Windows, use FormulaLocal
        If (Application.OperatingSystem Like "*Mac*") Then
            OSFormula = formVal
        Else
            Set sh = this.wksh
            prevForm = sh.Cells(1, 1).formula
            sh.Cells(1, 1).formula = "=" & formVal
            OSFormula = sh.Cells(1, 1).FormulaLocal
            sh.Cells(1, 1).formula = prevForm
        End If
    Else
        ' HList: always use FormulaLocal
        Set sh = this.wksh
        prevForm = sh.Cells(1, 1).formula
        sh.Cells(1, 1).formula = "=" & formVal
        OSFormula = sh.Cells(1, 1).FormulaLocal
        sh.Cells(1, 1).formula = prevForm
    End If
End Function

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a descriptive VBA error from this class
'@details
'Raises a custom VBA error with number vbObjectError + 516 and the
'class name as the source. Called by Create when required arguments are
'Nothing.
'@param errorMessage String. The error description to include in the raised error.
'@throws VBA.vbObjectError + 516 always.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, "Error : " & errorMessage
End Sub


'@section Checkings
'===============================================================================

Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingSuccess)
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("Variables")
    End If
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As IChecking)
    If collaboratorChecks Is Nothing Then Exit Sub
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("Variables")
    End If
    internalChecks.Append collaboratorChecks
    this.hasChecks = True
End Sub

Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

Public Property Get CheckingValues() As IChecking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property


'@section Interface Implementation
'@description Delegated members satisfying the IVarWriter contract.
'===============================================================================

Private Sub IVarWriter_WriteVariable(ByVal varName As String)
    WriteVariable varName
End Sub

Private Property Get IVarWriter_ValueOf(ByVal colName As String) As String
    IVarWriter_ValueOf = ValueOf(colName)
End Property

Private Property Get IVarWriter_HasCheckings() As Boolean
    IVarWriter_HasCheckings = HasCheckings
End Property

Private Property Get IVarWriter_CheckingValues() As IChecking
    Set IVarWriter_CheckingValues = CheckingValues
End Property
