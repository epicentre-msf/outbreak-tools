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

Private Type TVarWriter
    layer As Byte
    specs As ILinelistSpecs
    wksh As Worksheet
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
End Type

Private this As TVarWriter
Private Const CLASS_NAME As String = "VarWriter"
Private Const HLIST_START_ROW As Long = 8
Private Const VLIST_START_COL As Long = 5

'@section Factory
'===============================================================================

'@description Create a new VarWriter for the given layer and dependencies.
'@param layer VarWriterLayer enum value (HList=1, VList=2). Use Byte for macOS compatibility.
'@param specs The linelist specifications object providing dictionary, format, translations, etc.
'@param wksh The target main worksheet.
'@param printWksh Optional printed companion worksheet (HList only).
'@param crfWksh Optional CRF companion worksheet (HList only).
'@param dropdownObj Optional dropdown lists manager.
'@param customDropdownObj Optional custom dropdown lists manager (for choice_custom).
Public Function Create(ByVal layer As Byte, _
                       ByVal specs As ILinelistSpecs, _
                       ByVal wksh As Worksheet, _
                       Optional ByVal printWksh As Worksheet = Nothing, _
                       Optional ByVal crfWksh As Worksheet = Nothing, _
                       Optional ByVal dropdownObj As IDropdownLists = Nothing, _
                       Optional ByVal customDropdownObj As IDropdownLists = Nothing) As IVarWriter

    If specs Is Nothing Then ThrowError "Specifications object is required"
    If wksh Is Nothing Then ThrowError "Target worksheet is required"

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

Public Property Get Self() As IVarWriter
    Set Self = Me
End Property

'@section Internal setters (used by factory only)
'===============================================================================

Public Property Set InternalSpecs(ByVal specs As ILinelistSpecs)
    Set this.specs = specs
End Property

Public Property Set InternalWksh(ByVal wksh As Worksheet)
    Set this.wksh = wksh
End Property

Public Property Let InternalLayer(ByVal layer As Byte)
    this.layer = layer
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

Private Property Get SheetsObject() As ILLSheets
    If this.cachedSheets Is Nothing Then
        Set this.cachedSheets = LLSheets.Create(Dictionary())
    End If
    Set SheetsObject = this.cachedSheets
End Property

Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = this.specs.DesignFormat()
End Property

Private Property Get Trans() As ITranslationObject
    Set Trans = this.specs.TransObject()
End Property

Private Property Get FormulaDataObj() As IFormulaData
    Set FormulaDataObj = this.specs.FormulaDataObject()
End Property

'@section Value lookup
'===============================================================================

'@description Read a dictionary column value for the current variable.
Public Property Get ValueOf(ByVal colName As String) As String
    ValueOf = Variables().Value(colName:=colName, varName:=this.varName)
End Property

'@section Cell range resolution
'===============================================================================

'@description Get the main cell range for the current variable.
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

'@description Get the printed companion cell range (HList only).
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

'@description Get the CRF companion cell range (HList only).
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

'@section Main entry point
'===============================================================================

'@description Write one variable to the target worksheet(s).
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
End Sub

'@section Main sheet writing
'===============================================================================

'@description Write variable data to the main worksheet.
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
        AddGeoConditionalFormatting
        AddConditionalFormatting
        AddUniqueConditionalFormatting
    End If

    ' VList also has conditional formatting
    If this.layer = VarWriterLayerVList Then
        AddConditionalFormatting
    End If
End Sub

'@description Write variable label to the printed companion sheet (HList only).
Private Sub WritePrintSheet()
    If this.printWksh Is Nothing Then Exit Sub
    AddLabel usePrinted:=True
End Sub

'@description Write variable data to the CRF companion sheet (HList only).
Private Sub WriteCRFSheet()
    If this.crfWksh Is Nothing Then Exit Sub
    Dim crfRng As Range
    Set crfRng = VarCRFRange()
    If crfRng Is Nothing Then Exit Sub

    AddCRFLabel crfRng
    AddCRFType crfRng
    AddCRFChoices crfRng
End Sub

'@section AddLabel
'===============================================================================

'@description Write the variable label, sublabel, name, notes, control info, and format.
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

        ' Control info in hidden rows
        cellRng.Offset(-5).Value = varControl
        design.ApplyFormat cellRng.Offset(-5), scope:=LinelistHiddenCell
        cellRng.Offset(-6).Value = varListAutoOrigin
        design.ApplyFormat cellRng.Offset(-6), scope:=LinelistHiddenCell

        ' Geo labels
        If (InStr(1, varControl, "geo") > 0) Or (varControl = "hf") Or _
            (InStr(1, varName, "pcode_adm") > 0) Then
            design.ApplyFormat cellRng.Offset(-2), scope:=HListGeoHeader
            design.ApplyFormat cellRng, scope:=HListGeo
        End If

        ' Editable labels
        If ((varEditLab = "yes") Or (varControl = "custom")) Then cellRng.Offset(-2).Locked = False

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

        cellRng.Locked = False
        If (varEditLab = "yes") Then cellRng.Offset(, -1).Locked = False
        cellRng.Offset(, -2).Locked = True
        cellRng.Offset(, -3).Locked = True

        ' Control info in adjacent cell
        cellRng.Offset(, 1).Value = varControl
        design.ApplyFormat cellRng.Offset(, 1), scope:=LinelistHiddenCell
    End If
End Sub

'@section AddType
'===============================================================================

'@description Apply number format based on variable type and format.
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

'@section AddFormula
'===============================================================================

'@description Parse and apply formula/case_when/choice_formula to the variable cell.
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
End Sub

'@section AddChoices
'===============================================================================

'@description Set up dropdown validation for choice/list/geo variables.
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
    Dim labelOffset As Long

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

    ' Label offset depends on layer: HList = Offset(-2), VList = Offset(,-1)
    labelOffset = 0

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

'@section AddValidation
'===============================================================================

'@description Add min/max data validation to the variable cell.
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

'@section Conditional formatting
'===============================================================================

'@description Add expression-based conditional formatting.
Private Sub AddConditionalFormatting()
    Dim varValid As String
    Dim cellRng As Range
    Dim formatRng As Range
    Dim varObject As ILLVariables
    Dim shObj As ILLSheets
    Dim condAdd As String

    varValid = ValueOf("formatting condition")
    Set varObject = Variables()

    If (varValid = vbNullString) Then Exit Sub
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    Set formatRng = varObject.CellRange(colName:="formatting values", _
                                         varName:=this.varName)
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid)

    cellRng.FormatConditions.Add Type:=xlExpression, Formula1:="=(" & condAdd & "= 1)"
    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.Color = formatRng.Interior.Color
        .Font.Color = formatRng.Font.Color
        .Font.Bold = formatRng.Font.Bold
        .Font.Italic = formatRng.Font.Italic
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'@description Add unique/duplicate conditional formatting (HList only).
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

'@description Add geo pcode conditional formatting (HList only).
Private Sub AddGeoConditionalFormatting()
    Dim varValid As String
    Dim cellRng As Range
    Dim varObject As ILLVariables
    Dim shObj As ILLSheets
    Dim condAdd As String
    Dim varControl As String
    Dim actAdd As String

    varControl = ValueOf("control")
    If Not (InStr(1, varControl, "geo") > 1) Then Exit Sub

    varValid = "pcode" & "_" & this.varName
    Set varObject = Variables()
    If Not varObject.Contains(varValid) Then Exit Sub

    Set cellRng = VarRange()
    Set shObj = SheetsObject()
    condAdd = shObj.VariableAddress(varValid)
    actAdd = shObj.VariableAddress(this.varName)

    cellRng.FormatConditions.Add Type:=xlExpression, _
        Formula1:="= AND(" & condAdd & "= " & Chr(34) & Chr(34) & "," & actAdd & "<>" & Chr(34) & Chr(34) & ")"
    With cellRng.FormatConditions(cellRng.FormatConditions.Count)
        .Interior.Color = RGB(242, 147, 12)
        .SetFirstPriority
        .StopIfTrue = True
    End With
End Sub

'@section CRF companion writing
'===============================================================================

'@description Write CRF label for the current variable.
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

'@description Apply CRF-specific type formatting.
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

'@description Write CRF choice headers and values.
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

'@section Formula helpers
'===============================================================================

'@description Test if a formula fails when applied to a cell.
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

'@description Convert a formula value to an Excel-compatible format for validation.
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

    If Not varFormObject.Valid(formulaType:="linelist") Then Exit Function
    excelForm = varFormObject.ParsedLinelistFormula()
    If FailedFormula(excelForm) Then Exit Function

    If useOS Then
        excelForm = OSFormula(excelForm)
    Else
        excelForm = "= " & excelForm
    End If

    If Not varFormObject.HasLiterals Then
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
End Function

'@description Convert formula to OS-specific formula (handles macOS FormulaLocal).
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

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, "Error : " & errorMessage
End Sub

'@section Interface implementation
'===============================================================================

Private Sub IVarWriter_WriteVariable(ByVal varName As String)
    WriteVariable varName
End Sub

Private Property Get IVarWriter_ValueOf(ByVal colName As String) As String
    IVarWriter_ValueOf = ValueOf(colName)
End Property
