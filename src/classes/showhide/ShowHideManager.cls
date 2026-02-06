VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Consolidated show/hide visibility manager for all worksheet layers"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Consolidated show/hide visibility manager for all worksheet layers")
'@depends IShowHideManager, ILLdictionary, ILLVariables, LLVariables, BetterArray, IChecking

Option Explicit
Implements IShowHideManager


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "ShowHideManager"

' Persistence column headers
Private Const PERSIST_LAYER As String = "layer"
Private Const PERSIST_FIELD_KEY As String = "field_key"
Private Const PERSIST_HEADER As String = "header_text"
Private Const PERSIST_HIDDEN As String = "hidden_flag"

' Dictionary column names
Private Const COL_SHEET_NAME As String = "sheet name"
Private Const COL_SHEET_TYPE As String = "sheet type"
Private Const COL_MAIN_LABEL As String = "main label"
Private Const COL_SUB_SECTION As String = "sub section"
Private Const COL_CONTROL As String = "control"
Private Const COL_STATUS As String = "status"
Private Const COL_VISIBILITY As String = "visibility"
Private Const COL_REGISTER_BOOK As String = "register book"
Private Const COL_CRF_STATUS As String = "crf status"
Private Const COL_CRF_INDEX As String = "crf index"
Private Const COL_COLUMN_INDEX As String = "column index"
Private Const COL_LIST_AUTO As String = "list auto"


'@section State
'===============================================================================

Private Type TShowHideManager
    fieldKeys() As String
    headers() As String
    hidden() As Boolean
    mandatory() As Boolean
    dictHidden() As Boolean
    forceHidden() As Boolean
    positions() As Long
    controls() As String
    keyIndex As Collection
    entryCount As Long
    targetLayer As ShowHideWorksheetLayer
    sheetName As String
End Type

Private this As TShowHideManager


'@section Factory
'===============================================================================

' @description Build a ShowHideManager from dictionary data for a specific layer and sheet.
' @param dict Prepared or unprepared dictionary reference.
' @param layer Target worksheet layer.
' @param sheetName Clean sheet name (without print_/crf_ prefixes).
' @return IShowHideManager ready for use.
Public Function Create(ByVal dict As ILLdictionary, _
                       ByVal layer As ShowHideWorksheetLayer, _
                       ByVal sheetName As String) As IShowHideManager

    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ShowHideManager requires a dictionary reference"
    End If

    If LenB(Trim$(sheetName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "ShowHideManager requires a non-empty sheet name"
    End If

    ValidateLayer layer

    With New ShowHideManager
        .Initialise dict, layer, sheetName
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

' @description Return the current instance as an interface reference.
Public Property Get Self() As IShowHideManager
    Set Self = Me
End Property

' @description Number of entries in the manager.
Public Property Get EntryCount() As Long
    EntryCount = this.entryCount
End Property

' @description The worksheet layer this manager was built for.
Public Property Get TargetLayer() As ShowHideWorksheetLayer
    TargetLayer = this.targetLayer
End Property

' @description The sheet name used to filter dictionary variables during build.
Public Property Get SheetName() As String
    SheetName = this.sheetName
End Property


'@section Entry access
'===============================================================================

' @description Variable name at the given 1-based index.
Public Function FieldKey(ByVal index As Long) As String
    EnsureValidIndex index
    FieldKey = this.fieldKeys(index)
End Function

' @description Display header text at the given 1-based index.
Public Function HeaderText(ByVal index As Long) As String
    EnsureValidIndex index
    HeaderText = this.headers(index)
End Function

' @description Effective hidden state at the given 1-based index.
' Mandatory entries are always visible. Dictionary-hidden and force-hidden
' entries are always hidden. Otherwise returns the stored user choice.
Public Function IsHidden(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsHidden = EffectiveHidden(index)
End Function

' @description True when the entry cannot be hidden by the user.
Public Function IsMandatory(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsMandatory = this.mandatory(index)
End Function

' @description Column or row position in the target worksheet.
Public Function PositionIndex(ByVal index As Long) As Long
    EnsureValidIndex index
    PositionIndex = this.positions(index)
End Function

' @description Control type string (e.g. "geo1", "formula").
Public Function ControlType(ByVal index As Long) As String
    EnsureValidIndex index
    ControlType = this.controls(index)
End Function


'@section Lookup
'===============================================================================

' @description Find the 1-based index of a field key. Returns 0 when absent.
Public Function IndexOf(ByVal fieldKey As String) As Long
    Dim key As String
    Dim result As Long

    key = LCase$(Trim$(fieldKey))
    If LenB(key) = 0 Then
        IndexOf = 0
        Exit Function
    End If

    If this.keyIndex Is Nothing Then
        IndexOf = 0
        Exit Function
    End If

    result = 0
    On Error Resume Next
        result = this.keyIndex.Item(key)
    On Error GoTo 0

    IndexOf = result
End Function

' @description True when the field key exists in the manager.
Public Function HasField(ByVal fieldKey As String) As Boolean
    HasField = (IndexOf(fieldKey) > 0)
End Function


'@section Visibility management
'===============================================================================

' @description Set hidden state for the entry at the given index.
' Mandatory entries are silently ignored.
Public Sub SetHidden(ByVal index As Long, ByVal hidden As Boolean)
    EnsureValidIndex index

    ' Mandatory entries cannot be hidden
    If this.mandatory(index) Then Exit Sub

    this.hidden(index) = hidden
End Sub

' @description Bulk show or hide all non-mandatory, non-locked entries.
Public Sub SetAllOptionalHidden(ByVal hidden As Boolean)
    Dim idx As Long

    For idx = 1 To this.entryCount
        If Not this.mandatory(idx) And _
           Not this.dictHidden(idx) And _
           Not this.forceHidden(idx) Then
            this.hidden(idx) = hidden
        End If
    Next idx
End Sub


'@section Plan generation
'===============================================================================

' @description Return visibility data for all entries.
' Each element is Array(headerText, fieldKey, effectiveHidden).
Public Function Plan() As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.LowerBound = 1

    For idx = 1 To this.entryCount
        result.Push Array(this.headers(idx), this.fieldKeys(idx), EffectiveHidden(idx))
    Next idx

    Set Plan = result
End Function


'@section Persistence
'===============================================================================

' @description Export current visibility state to a ListObject.
' Clears existing rows for this layer before writing.
Public Sub ExportPlan(ByVal lo As ListObject)
    Dim token As String
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim headerCol As Long
    Dim hiddenCol As Long
    Dim idx As Long
    Dim newRow As ListRow

    If lo Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Cannot export to an empty ListObject"
    End If

    token = LayerToken(this.targetLayer)
    EnsurePersistColumns lo

    layerCol = PersistColumnIndex(lo, PERSIST_LAYER)
    fieldCol = PersistColumnIndex(lo, PERSIST_FIELD_KEY)
    headerCol = PersistColumnIndex(lo, PERSIST_HEADER)
    hiddenCol = PersistColumnIndex(lo, PERSIST_HIDDEN)

    ' Clear existing rows for this layer
    ClearLayerRows lo, token, layerCol

    ' Write current entries
    For idx = 1 To this.entryCount
        Set newRow = lo.ListRows.Add
        With newRow.Range
            .Cells(1, layerCol).Value = token
            .Cells(1, fieldCol).Value = this.fieldKeys(idx)
            .Cells(1, headerCol).Value = this.headers(idx)
            .Cells(1, hiddenCol).Value = BoolToString(EffectiveHidden(idx))
        End With
    Next idx
End Sub

' @description Import persisted visibility, updating hidden states for matching entries.
Public Sub ImportPlan(ByVal lo As ListObject)
    Dim token As String
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim hiddenCol As Long
    Dim currentRow As ListRow
    Dim rowLayer As String
    Dim rowFieldKey As String
    Dim rowHidden As Boolean
    Dim entryIndex As Long

    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    token = LayerToken(this.targetLayer)
    layerCol = PersistColumnIndex(lo, PERSIST_LAYER)
    fieldCol = PersistColumnIndex(lo, PERSIST_FIELD_KEY)
    hiddenCol = PersistColumnIndex(lo, PERSIST_HIDDEN)

    ' If any required column is missing, silently exit
    If layerCol = 0 Or fieldCol = 0 Or hiddenCol = 0 Then Exit Sub

    For Each currentRow In lo.ListRows
        rowLayer = CStr(currentRow.Range.Cells(1, layerCol).Value)
        If StrComp(rowLayer, token, vbTextCompare) = 0 Then
            rowFieldKey = Trim$(CStr(currentRow.Range.Cells(1, fieldCol).Value))
            rowHidden = ToBool(currentRow.Range.Cells(1, hiddenCol).Value)

            entryIndex = IndexOf(rowFieldKey)
            If entryIndex > 0 Then
                ' Only update non-mandatory entries
                If Not this.mandatory(entryIndex) Then
                    this.hidden(entryIndex) = rowHidden
                End If
            End If
        End If
    Next currentRow
End Sub


'@section Initialisation
'===============================================================================

' @description Prime internal state from dictionary data.
' Called by the Create factory after validation.
Public Sub Initialise(ByVal dict As ILLdictionary, _
                      ByVal layer As ShowHideWorksheetLayer, _
                      ByVal sheetName As String)

    Dim vars As ILLVariables
    Dim varNames As BetterArray
    Dim idx As Long
    Dim varName As String
    Dim count As Long
    Dim maxEntries As Long
    Dim hasVisibility As Boolean
    Dim hasStatus As Boolean
    Dim hasRegisterBook As Boolean
    Dim hasCrfStatus As Boolean
    Dim hasListAuto As Boolean
    Dim hasSubSection As Boolean
    Dim hasSheetType As Boolean
    Dim hasCrfIndex As Boolean
    Dim hasColumnIndex As Boolean
    Dim headerVal As String
    Dim controlVal As String
    Dim subSectionVal As String
    Dim posValue As String
    Dim visValue As String
    Dim loweredVis As String
    Dim isMandatoryEntry As Boolean
    Dim isForceHiddenEntry As Boolean
    Dim listAutoVal As String

    this.targetLayer = layer
    this.sheetName = Trim$(sheetName)
    this.entryCount = 0
    Set this.keyIndex = New Collection

    If Not dict.Prepared Then dict.Prepare

    Set vars = LLVariables.Create(dict)
    Set varNames = vars.VariableNames

    If varNames Is Nothing Then Exit Sub
    If varNames.Length = 0 Then Exit Sub

    ' Cache column availability
    hasVisibility = dict.ColumnExists(COL_VISIBILITY)
    hasStatus = dict.ColumnExists(COL_STATUS)
    hasRegisterBook = dict.ColumnExists(COL_REGISTER_BOOK)
    hasCrfStatus = dict.ColumnExists(COL_CRF_STATUS)
    hasListAuto = dict.ColumnExists(COL_LIST_AUTO)
    hasSubSection = dict.ColumnExists(COL_SUB_SECTION)
    hasSheetType = dict.ColumnExists(COL_SHEET_TYPE)
    hasCrfIndex = dict.ColumnExists(COL_CRF_INDEX)
    hasColumnIndex = dict.ColumnExists(COL_COLUMN_INDEX)

    ' Pre-allocate arrays (worst case: all vars match)
    maxEntries = varNames.Length
    ReDim this.fieldKeys(1 To maxEntries)
    ReDim this.headers(1 To maxEntries)
    ReDim this.hidden(1 To maxEntries)
    ReDim this.mandatory(1 To maxEntries)
    ReDim this.dictHidden(1 To maxEntries)
    ReDim this.forceHidden(1 To maxEntries)
    ReDim this.positions(1 To maxEntries)
    ReDim this.controls(1 To maxEntries)

    count = 0

    For idx = varNames.LowerBound To varNames.UpperBound
        varName = CStr(varNames.Item(idx))
        If LenB(varName) = 0 Then GoTo NextVar

        ' Match sheet name
        If Not MatchesSheet(vars, varName) Then GoTo NextVar

        ' Match sheet type to target layer (safety filter)
        If hasSheetType Then
            If Not MatchesLayer(Normalise(SafeValue(vars, COL_SHEET_TYPE, varName))) Then GoTo NextVar
        End If

        ' Guard against duplicate field keys
        If HasField(varName) Then GoTo NextVar

        count = count + 1

        ' Field key
        this.fieldKeys(count) = varName

        ' Control type (normalised for comparisons)
        controlVal = Normalise(SafeValue(vars, COL_CONTROL, varName))
        this.controls(count) = controlVal

        ' Header text with geo/hf prefix
        headerVal = SafeValue(vars, COL_MAIN_LABEL, varName)
        If LenB(headerVal) = 0 Then headerVal = varName

        If hasSubSection Then
            If InStr(1, controlVal, "geo") > 0 Or InStr(1, controlVal, "hf") > 0 Then
                subSectionVal = SafeValue(vars, COL_SUB_SECTION, varName)
                If LenB(subSectionVal) > 0 Then
                    headerVal = subSectionVal & ": " & headerVal
                End If
            End If
        End If
        this.headers(count) = headerVal

        ' Position index (column or row position in the worksheet)
        posValue = vbNullString
        If layer = ShowHideLayerCRF And hasCrfIndex Then
            posValue = SafeValue(vars, COL_CRF_INDEX, varName)
        ElseIf hasColumnIndex Then
            posValue = SafeValue(vars, COL_COLUMN_INDEX, varName)
        End If
        If IsNumeric(posValue) And LenB(posValue) > 0 Then
            this.positions(count) = CLng(posValue)
        Else
            this.positions(count) = 0
        End If

        ' Determine visibility status from the appropriate dictionary column
        visValue = ResolveVisibilityStatus(vars, varName, hasVisibility, _
                                           hasStatus, hasRegisterBook, hasCrfStatus)
        loweredVis = Normalise(visValue)

        ' Mandatory flag
        isMandatoryEntry = (loweredVis = "always visible") Or (loweredVis = "mandatory")
        If layer = ShowHideLayerHList And controlVal = "geo1" Then
            isMandatoryEntry = True
        End If
        this.mandatory(count) = isMandatoryEntry

        ' Dictionary hidden flag (permanently hidden, user cannot toggle)
        this.dictHidden(count) = (loweredVis = "always hidden")

        ' Force hidden flag (formula/choice_formula/case_when/list_auto on CRF or Printed)
        isForceHiddenEntry = False
        If layer = ShowHideLayerCRF Or layer = ShowHideLayerPrinted Then
            If controlVal = "formula" Or controlVal = "choice_formula" _
               Or controlVal = "case_when" Then
                isForceHiddenEntry = True
            End If
            If hasListAuto Then
                listAutoVal = Normalise(SafeValue(vars, COL_LIST_AUTO, varName))
                If listAutoVal = "yes" Then isForceHiddenEntry = True
            End If
        End If
        this.forceHidden(count) = isForceHiddenEntry

        ' Initial hidden state
        this.hidden(count) = DetermineInitialHidden(loweredVis, layer)

        ' Add to key index (lowercase for case-insensitive lookup)
        this.keyIndex.Add count, LCase$(varName)

NextVar:
    Next idx

    ' Trim arrays to actual count
    this.entryCount = count
    If count > 0 And count < maxEntries Then
        ReDim Preserve this.fieldKeys(1 To count)
        ReDim Preserve this.headers(1 To count)
        ReDim Preserve this.hidden(1 To count)
        ReDim Preserve this.mandatory(1 To count)
        ReDim Preserve this.dictHidden(1 To count)
        ReDim Preserve this.forceHidden(1 To count)
        ReDim Preserve this.positions(1 To count)
        ReDim Preserve this.controls(1 To count)
    End If
End Sub


'@section Private helpers
'===============================================================================

' @description Compute effective hidden state for an entry.
Private Function EffectiveHidden(ByVal index As Long) As Boolean
    If this.mandatory(index) Then
        EffectiveHidden = False
    ElseIf this.dictHidden(index) Then
        EffectiveHidden = True
    ElseIf this.forceHidden(index) Then
        EffectiveHidden = True
    Else
        EffectiveHidden = this.hidden(index)
    End If
End Function

' @description Check if variable's sheet name matches the target sheet.
Private Function MatchesSheet(ByVal vars As ILLVariables, ByVal varName As String) As Boolean
    Dim varSheetName As String

    varSheetName = Normalise(SafeValue(vars, COL_SHEET_NAME, varName))
    MatchesSheet = (StrComp(varSheetName, Normalise(this.sheetName), vbTextCompare) = 0)
End Function

' @description Check if a sheet type string matches the target layer.
Private Function MatchesLayer(ByVal sheetType As String) As Boolean
    Select Case this.targetLayer
        Case ShowHideLayerHList, ShowHideLayerPrinted
            MatchesLayer = (InStr(1, sheetType, "hlist", vbBinaryCompare) > 0)
        Case ShowHideLayerVList
            MatchesLayer = (InStr(1, sheetType, "vlist", vbBinaryCompare) > 0)
        Case ShowHideLayerCRF
            MatchesLayer = (InStr(1, sheetType, "crf", vbBinaryCompare) > 0)
        Case Else
            MatchesLayer = False
    End Select
End Function

' @description Resolve which dictionary column to use for visibility status.
Private Function ResolveVisibilityStatus(ByVal vars As ILLVariables, _
                                          ByVal varName As String, _
                                          ByVal hasVisibility As Boolean, _
                                          ByVal hasStatus As Boolean, _
                                          ByVal hasRegisterBook As Boolean, _
                                          ByVal hasCrfStatus As Boolean) As String
    Select Case this.targetLayer
        Case ShowHideLayerPrinted
            If hasRegisterBook Then
                ResolveVisibilityStatus = SafeValue(vars, COL_REGISTER_BOOK, varName)
            Else
                ResolveVisibilityStatus = vbNullString
            End If
        Case ShowHideLayerCRF
            If hasCrfStatus Then
                ResolveVisibilityStatus = SafeValue(vars, COL_CRF_STATUS, varName)
            Else
                ResolveVisibilityStatus = vbNullString
            End If
        Case Else
            If hasVisibility Then
                ResolveVisibilityStatus = SafeValue(vars, COL_VISIBILITY, varName)
            ElseIf hasStatus Then
                ResolveVisibilityStatus = SafeValue(vars, COL_STATUS, varName)
            Else
                ResolveVisibilityStatus = vbNullString
            End If
    End Select
End Function

' @description Determine initial hidden state from a visibility status string.
Private Function DetermineInitialHidden(ByVal loweredStatus As String, _
                                         ByVal layer As ShowHideWorksheetLayer) As Boolean
    Select Case loweredStatus
        Case "hidden", "hidden on creation", "always hidden", "optional, hidden", "do not print"
            DetermineInitialHidden = True
        Case Else
            If layer = ShowHideLayerPrinted Then
                DetermineInitialHidden = (InStr(1, loweredStatus, "hidden") > 0)
            Else
                DetermineInitialHidden = False
            End If
    End Select
End Function

' @description Safely read a variable value from dictionary, returning empty string on error.
Private Function SafeValue(ByVal vars As ILLVariables, _
                            ByVal colName As String, _
                            ByVal varName As String) As String
    Dim result As String

    On Error Resume Next
        result = vars.Value(colName, varName)
    On Error GoTo 0

    SafeValue = result
End Function

' @description Trim and lowercase a string for normalised comparisons.
Private Function Normalise(ByVal value As String) As String
    Normalise = LCase$(Trim$(value))
End Function

' @description Validate that an index is within bounds.
Private Sub EnsureValidIndex(ByVal index As Long)
    If index < 1 Or index > this.entryCount Then
        ThrowError ProjectError.InvalidArgument, _
                   "Entry index " & CStr(index) & " is outside range 1.." & CStr(this.entryCount)
    End If
End Sub

' @description Validate layer enum value.
Private Sub ValidateLayer(ByVal layer As ShowHideWorksheetLayer)
    If layer < ShowHideLayerHList Or layer > ShowHideLayerCRF Then
        ThrowError ProjectError.InvalidArgument, _
                   "ShowHideManager layer value is outside expected range"
    End If
End Sub

' @description Map a layer enum to a persistence token string.
Private Function LayerToken(ByVal layer As ShowHideWorksheetLayer) As String
    Select Case layer
        Case ShowHideLayerHList
            LayerToken = "hlist"
        Case ShowHideLayerPrinted
            LayerToken = "printed"
        Case ShowHideLayerVList
            LayerToken = "vlist"
        Case ShowHideLayerCRF
            LayerToken = "crf"
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unsupported layer value: " & CStr(layer)
    End Select
End Function

' @description Find a column index in a ListObject header row, returning 0 if absent.
Private Function PersistColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim found As Range

    If lo Is Nothing Then
        PersistColumnIndex = 0
        Exit Function
    End If

    On Error Resume Next
        Set found = lo.HeaderRowRange.Find(What:=headerName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If found Is Nothing Then
        PersistColumnIndex = 0
    Else
        PersistColumnIndex = found.Column - lo.HeaderRowRange.Column + 1
    End If
End Function

' @description Ensure the ListObject has all required persistence columns.
Private Sub EnsurePersistColumns(ByVal lo As ListObject)
    Dim required As Variant
    Dim idx As Long

    required = Array(PERSIST_LAYER, PERSIST_FIELD_KEY, PERSIST_HEADER, PERSIST_HIDDEN)
    For idx = LBound(required) To UBound(required)
        If PersistColumnIndex(lo, CStr(required(idx))) = 0 Then
            ThrowError ProjectError.ElementNotFound, _
                       "Persistence table is missing column: " & CStr(required(idx))
        End If
    Next idx
End Sub

' @description Remove all rows matching the layer token.
Private Sub ClearLayerRows(ByVal lo As ListObject, ByVal token As String, ByVal layerCol As Long)
    Dim idx As Long

    If lo.ListRows.Count = 0 Then Exit Sub

    For idx = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(lo.ListRows(idx).Range.Cells(1, layerCol).Value), _
                   token, vbTextCompare) = 0 Then
            lo.ListRows(idx).Delete
        End If
    Next idx
End Sub

' @description Convert Boolean to persistence string.
Private Function BoolToString(ByVal value As Boolean) As String
    If value Then
        BoolToString = "true"
    Else
        BoolToString = "false"
    End If
End Function

' @description Parse a persisted value to Boolean.
Private Function ToBool(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        ToBool = CBool(value)
    ElseIf VarType(value) = vbString Then
        Select Case LCase$(Trim$(CStr(value)))
            Case "true", "yes", "1"
                ToBool = True
            Case Else
                ToBool = False
        End Select
    ElseIf IsNumeric(value) Then
        ToBool = (CLng(value) <> 0)
    Else
        ToBool = False
    End If
End Function

' @description Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================

Private Property Get IShowHideManager_EntryCount() As Long
    IShowHideManager_EntryCount = EntryCount
End Property

Private Property Get IShowHideManager_TargetLayer() As ShowHideWorksheetLayer
    IShowHideManager_TargetLayer = TargetLayer
End Property

Private Property Get IShowHideManager_SheetName() As String
    IShowHideManager_SheetName = SheetName
End Property

Private Function IShowHideManager_FieldKey(ByVal index As Long) As String
    IShowHideManager_FieldKey = FieldKey(index)
End Function

Private Function IShowHideManager_HeaderText(ByVal index As Long) As String
    IShowHideManager_HeaderText = HeaderText(index)
End Function

Private Function IShowHideManager_IsHidden(ByVal index As Long) As Boolean
    IShowHideManager_IsHidden = IsHidden(index)
End Function

Private Function IShowHideManager_IsMandatory(ByVal index As Long) As Boolean
    IShowHideManager_IsMandatory = IsMandatory(index)
End Function

Private Function IShowHideManager_PositionIndex(ByVal index As Long) As Long
    IShowHideManager_PositionIndex = PositionIndex(index)
End Function

Private Function IShowHideManager_ControlType(ByVal index As Long) As String
    IShowHideManager_ControlType = ControlType(index)
End Function

Private Function IShowHideManager_IndexOf(ByVal fieldKey As String) As Long
    IShowHideManager_IndexOf = IndexOf(fieldKey)
End Function

Private Function IShowHideManager_HasField(ByVal fieldKey As String) As Boolean
    IShowHideManager_HasField = HasField(fieldKey)
End Function

Private Sub IShowHideManager_SetHidden(ByVal index As Long, ByVal hidden As Boolean)
    SetHidden index, hidden
End Sub

Private Sub IShowHideManager_SetAllOptionalHidden(ByVal hidden As Boolean)
    SetAllOptionalHidden hidden
End Sub

Private Function IShowHideManager_Plan() As BetterArray
    Set IShowHideManager_Plan = Plan()
End Function

Private Sub IShowHideManager_ExportPlan(ByVal lo As ListObject)
    ExportPlan lo
End Sub

Private Sub IShowHideManager_ImportPlan(ByVal lo As ListObject)
    ImportPlan lo
End Sub
