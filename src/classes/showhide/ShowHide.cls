VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Which variables of one worksheet layer the user can see"

'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Which variables of one worksheet layer the user can see")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'The four worksheet layers a linelist can show or hide. The enum lives on this
'class, so naming ShowHide brings the values with it.
Public Enum ShowHideWorksheetLayer
    ShowHideLayerHList = 1
    ShowHideLayerPrinted = 2
    ShowHideLayerVList = 3
    ShowHideLayerCRF = 4
End Enum


'@class ShowHide
'@description
'Holds the list of variables one worksheet layer offers the user, and the
'visible or hidden choice for each of them. The list is built from the
'dictionary. EntryCount, FieldKey, HeaderText, IsHidden, PositionIndex,
'ControlType and AuthoredVertical read one entry, IndexOf and HasField find
'one, SetHidden and SetAllOptionalHidden change the choices, Apply writes them
'through a layout and Adopt reads them back from a saved set. The class reads
'the dictionary alone: the worksheet is ShowHideLayout's work and the saved
'choices are ShowHideStore's.
'
'FOUR KINDS OF ENTRY
'-------------------------------------------------------------------------------
'A variable the dictionary marks as hidden is EXCLUDED. It never enters the
'list, because the designer said the user must never see it.
'
'Of the variables that do enter, each is one of three kinds:
'
'  mandatory  listed, always visible, the user cannot change it.
'  locked     listed, always hidden, the user cannot change it. A calculated
'             column on a printed sheet or on a CRF lands here.
'  free       listed, starts where the dictionary says, follows the user.
'
'IsFree is the single question a click handler asks before it changes anything.
'
'WHICH DICTIONARY COLUMN SAYS WHAT
'-------------------------------------------------------------------------------
'  HList, VList   visibility
'  Printed        register book
'  CRF            crf status
'
'The position of a variable comes from `column index`, and from `crf index` on
'the CRF layer. A variable with no position is listed, and the layout leaves it
'where it is.
'
'THE STATE IS A BOOLEAN
'-------------------------------------------------------------------------------
'Hidden is True or False. The display words the form shows ("Shown", "Hidden",
'"Mandatory") are translated at the form and stored nowhere. A user who changes
'the interface language between two sessions keeps every choice.
'@depends LLdictionary, LLVariables, ShowHideLayout, BetterArray, Checking
'@version 1.0 (2026-08-01)


Private Const CLASS_NAME As String = "ShowHide"

'Dictionary column names
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

'The one `register book` value that asks for a turned header
Private Const PRINT_VERTICAL As String = "print, vertical header"

Private Type TShowHide
    fieldKeys() As String
    headers() As String
    hidden() As Boolean
    mandatory() As Boolean
    locked() As Boolean
    verticalHeader() As Boolean
    positions() As Long
    controls() As String
    keyIndex As Collection
    entryCount As Long
    targetLayer As Byte
    sheetName As String
End Type

Private this As TShowHide


'@section Instantiation
'===============================================================================
'@description The factory and the build it runs.

'@label:create
'@sub-title Build the entry list of one layer of one sheet
'@details
'Walks the dictionary once, keeps the variables of the named sheet whose sheet
'type suits the layer, and files each one as mandatory, locked or free. The
'dictionary is prepared when it arrives unprepared, because the four columns
'this class reads (visibility, crf status, column index, crf index) are all
'written by Prepare.
'@param dict LLdictionary. The dictionary to read.
'@param layer Byte. A ShowHideWorksheetLayer value. Which of the four layers to build.
'@param sheetName String. The base sheet name, with no print_ or crf_ prefix.
'@return ShowHide. The entry list, ready to read.
'@throws ProjectError.ObjectNotInitialized When the dictionary is Nothing.
'@throws ProjectError.InvalidArgument When the sheet name is empty or the layer
'  is outside the four values.
'@export
Public Function Create(ByVal dict As LLdictionary, _
                       ByVal layer As Byte, _
                       ByVal sheetName As String) As ShowHide

    Dim instance As ShowHide

    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A dictionary is needed to know which variables a sheet offers"
    End If

    If LenB(Trim$(sheetName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "A sheet name is needed to build the show/hide entries"
    End If

    ValidateLayer layer

    Set instance = New ShowHide
    instance.Initialise dict, layer, sheetName
    Set Create = instance
End Function


'@section Accessors
'===============================================================================
'@description What the list holds, and what each entry is.

'@label:entry-count
'@prop-title The number of entries the user is offered
'@export
Public Property Get EntryCount() As Long
    EntryCount = this.entryCount
End Property

'@label:target-layer
'@prop-title The layer this list was built for
'@export
Public Property Get TargetLayer() As Byte
    TargetLayer = this.targetLayer
End Property

'@label:sheet-name
'@prop-title The base sheet name the entries belong to
'@export
Public Property Get SheetName() As String
    SheetName = this.sheetName
End Property

'@label:field-key
'@sub-title The variable name of one entry
'@param index Long. A one based entry index.
'@return String. The variable name as the dictionary spells it.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function FieldKey(ByVal index As Long) As String
    EnsureValidIndex index
    FieldKey = this.fieldKeys(index)
End Function

'@label:header-text
'@sub-title The label the form shows for one entry
'@param index Long. A one based entry index.
'@return String. The main label, carrying its sub section for a geo or hf
'  variable.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function HeaderText(ByVal index As Long) As String
    EnsureValidIndex index
    HeaderText = this.headers(index)
End Function

'@label:is-hidden
'@sub-title Whether one entry is hidden right now
'@details
'A mandatory entry answers False and a locked entry answers True, whatever the
'user asked for. Every other entry answers the stored choice.
'@param index Long. A one based entry index.
'@return Boolean. True when the variable is hidden.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function IsHidden(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsHidden = EffectiveHidden(index)
End Function

'@label:is-mandatory
'@sub-title Whether one entry is always visible
'@param index Long. A one based entry index.
'@return Boolean. True when the user cannot hide the variable.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function IsMandatory(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsMandatory = this.mandatory(index)
End Function

'@label:is-locked
'@sub-title Whether one entry is always hidden
'@param index Long. A one based entry index.
'@return Boolean. True when the variable is held hidden and the user cannot
'  show it.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function IsLocked(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsLocked = this.locked(index)
End Function

'@label:is-free
'@sub-title Whether the user may change one entry
'@details
'The single guard a click handler asks. An entry that is neither mandatory nor
'locked follows whatever the user picks.
'@param index Long. A one based entry index.
'@return Boolean. True when the entry follows the user.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function IsFree(ByVal index As Long) As Boolean
    EnsureValidIndex index
    IsFree = Not (this.mandatory(index) Or this.locked(index))
End Function

'@label:position-index
'@sub-title Where one entry sits on the worksheet
'@param index Long. A one based entry index.
'@return Long. The column or row number, or 0 when the dictionary gives none.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function PositionIndex(ByVal index As Long) As Long
    EnsureValidIndex index
    PositionIndex = this.positions(index)
End Function

'@label:control-type
'@sub-title The dictionary control of one entry
'@param index Long. A one based entry index.
'@return String. The control value in lower case, such as "geo1" or "formula".
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function ControlType(ByVal index As Long) As String
    EnsureValidIndex index
    ControlType = this.controls(index)
End Function

'@label:authored-vertical
'@sub-title The header direction the dictionary asked for
'@details
'Read only, and true for a variable whose `register book` reads
'"print, vertical header". A reset restores this direction without the entry
'list ever owning the live orientation of the sheet.
'@param index Long. A one based entry index.
'@return Boolean. True when the printed header should be turned.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function AuthoredVertical(ByVal index As Long) As Boolean
    EnsureValidIndex index
    AuthoredVertical = this.verticalHeader(index)
End Function


'@section Lookup
'===============================================================================
'@description Finding an entry by its variable name.

'@label:index-of
'@sub-title Find an entry by variable name
'@details
'The name is trimmed and lower cased on the way in and on the way out, so a
'dictionary name carrying a stray space is still reachable.
'@param fieldKey String. The variable name to look for.
'@return Long. The one based entry index, or 0 when the name is absent.
'@export
Public Function IndexOf(ByVal fieldKey As String) As Long
    Dim key As String
    Dim result As Long

    key = NormaliseKey(fieldKey)
    If LenB(key) = 0 Then Exit Function
    If this.keyIndex Is Nothing Then Exit Function

    result = 0
    On Error Resume Next
        result = this.keyIndex.Item(key)
    On Error GoTo 0

    IndexOf = result
End Function

'@label:has-field
'@sub-title Whether a variable name is in the list
'@param fieldKey String. The variable name to look for.
'@return Boolean. True when the list carries the name.
'@export
Public Function HasField(ByVal fieldKey As String) As Boolean
    HasField = (IndexOf(fieldKey) > 0)
End Function


'@section Choices
'===============================================================================
'@description Changing what the user sees.

'@label:set-hidden
'@sub-title Hide or show one entry
'@details
'A mandatory or locked entry keeps its state and the call returns quietly, so a
'caller that loops the whole list needs no guard of its own.
'@param index Long. A one based entry index.
'@param hidden Boolean. True to hide the variable.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Sub SetHidden(ByVal index As Long, ByVal hidden As Boolean)
    EnsureValidIndex index
    If this.mandatory(index) Or this.locked(index) Then Exit Sub
    this.hidden(index) = hidden
End Sub

'@label:set-all-optional-hidden
'@sub-title Hide or show every entry the user owns
'@param hidden Boolean. True to hide all the free entries.
'@export
Public Sub SetAllOptionalHidden(ByVal hidden As Boolean)
    Dim idx As Long

    For idx = 1 To this.entryCount
        If Not (this.mandatory(idx) Or this.locked(idx)) Then
            this.hidden(idx) = hidden
        End If
    Next idx
End Sub


'@section Worksheet
'===============================================================================
'@description The two calls that put the list and a sheet in step.

'@label:apply
'@sub-title Put the worksheet in the state the list describes
'@details
'The one place the entry list reaches a sheet, and it goes through the layout,
'so the row versus column question and the protection bracket are answered in
'one place. Entries with no position are skipped.
'@param layout ShowHideLayout. The sheet to reconcile.
'@return Long. The number of positions that were set.
'@export
Public Function Apply(ByVal layout As ShowHideLayout) As Long
    Dim idx As Long
    Dim position As Long
    Dim applied As Long

    If layout Is Nothing Then Exit Function

    layout.BeginBatch
    For idx = 1 To this.entryCount
        position = this.positions(idx)
        If position > 0 Then
            layout.SetHidden position, EffectiveHidden(idx)
            applied = applied + 1
        End If
    Next idx
    layout.EndBatch

    Apply = applied
End Function

'@label:adopt
'@sub-title Read the worksheet back into the list
'@details
'For the case where the user hid a column by hand rather than through the form.
'Mandatory and locked entries keep what the dictionary said.
'@param layout ShowHideLayout. The sheet to read.
'@return Long. The number of entries that were read.
'@export
Public Function Adopt(ByVal layout As ShowHideLayout) As Long
    Dim idx As Long
    Dim position As Long
    Dim adopted As Long

    If layout Is Nothing Then Exit Function

    For idx = 1 To this.entryCount
        position = this.positions(idx)
        If position > 0 And Not (this.mandatory(idx) Or this.locked(idx)) Then
            this.hidden(idx) = layout.IsHidden(position)
            adopted = adopted + 1
        End If
    Next idx

    Adopt = adopted
End Function


'@section Build
'===============================================================================
'@description The dictionary walk the factory runs.

'@label:initialise
'@sub-title Fill the entry list from the dictionary
'@details
'Public because the factory calls it on a fresh instance. Every variable of the
'sheet is read once and the arrays are trimmed to what matched.
'@param dict LLdictionary. The dictionary to read.
'@param layer Byte. A ShowHideWorksheetLayer value. The layer being built.
'@param sheetName String. The base sheet name.
'@export
Public Sub Initialise(ByVal dict As LLdictionary, _
                      ByVal layer As Byte, _
                      ByVal sheetName As String)

    Dim vars As LLVariables
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
    Dim layerStatus As String
    Dim registerBook As String

    this.targetLayer = layer
    this.sheetName = Trim$(sheetName)
    this.entryCount = 0
    Set this.keyIndex = New Collection

    If Not dict.Prepared Then dict.Prepare

    Set vars = LLVariables.Create(dict)
    Set varNames = vars.VariableNames

    If varNames Is Nothing Then Exit Sub
    If varNames.Length = 0 Then Exit Sub

    hasVisibility = dict.ColumnExists(COL_VISIBILITY)
    hasStatus = dict.ColumnExists(COL_STATUS)
    hasRegisterBook = dict.ColumnExists(COL_REGISTER_BOOK)
    hasCrfStatus = dict.ColumnExists(COL_CRF_STATUS)
    hasListAuto = dict.ColumnExists(COL_LIST_AUTO)
    hasSubSection = dict.ColumnExists(COL_SUB_SECTION)
    hasSheetType = dict.ColumnExists(COL_SHEET_TYPE)
    hasCrfIndex = dict.ColumnExists(COL_CRF_INDEX)
    hasColumnIndex = dict.ColumnExists(COL_COLUMN_INDEX)

    maxEntries = varNames.Length
    ReDim this.fieldKeys(1 To maxEntries)
    ReDim this.headers(1 To maxEntries)
    ReDim this.hidden(1 To maxEntries)
    ReDim this.mandatory(1 To maxEntries)
    ReDim this.locked(1 To maxEntries)
    ReDim this.verticalHeader(1 To maxEntries)
    ReDim this.positions(1 To maxEntries)
    ReDim this.controls(1 To maxEntries)

    count = 0

    For idx = varNames.LowerBound To varNames.UpperBound
        varName = CStr(varNames.Item(idx))
        If LenB(Trim$(varName)) = 0 Then GoTo NextVar

        If Not MatchesSheet(vars, varName) Then GoTo NextVar

        If hasSheetType Then
            If Not MatchesLayer(Normalise(SafeValue(vars, COL_SHEET_TYPE, varName))) Then GoTo NextVar
        End If

        'A variable the designer hid never reaches the user's list
        If hasVisibility Then
            If Normalise(SafeValue(vars, COL_VISIBILITY, varName)) = "always hidden" Then GoTo NextVar
        End If

        If HasField(varName) Then GoTo NextVar

        count = count + 1

        this.fieldKeys(count) = varName

        controlVal = Normalise(SafeValue(vars, COL_CONTROL, varName))
        this.controls(count) = controlVal

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

        posValue = vbNullString
        If layer = ShowHideLayerCRF And hasCrfIndex Then
            posValue = SafeValue(vars, COL_CRF_INDEX, varName)
        ElseIf hasColumnIndex Then
            posValue = SafeValue(vars, COL_COLUMN_INDEX, varName)
        End If
        If LenB(posValue) > 0 And IsNumeric(posValue) Then
            this.positions(count) = CLng(posValue)
        Else
            this.positions(count) = 0
        End If

        registerBook = vbNullString
        If hasRegisterBook Then registerBook = Normalise(SafeValue(vars, COL_REGISTER_BOOK, varName))
        this.verticalHeader(count) = (registerBook = PRINT_VERTICAL)

        layerStatus = Normalise(StatusForLayer(vars, varName, registerBook, _
                                            hasVisibility, hasStatus, hasCrfStatus))

        this.mandatory(count) = IsMandatoryStatus(layerStatus, controlVal, layer)
        this.locked(count) = IsLockedStatus(layerStatus, controlVal, layer, _
                                            this.mandatory(count), hasListAuto, vars, varName)
        this.hidden(count) = InitialHidden(layerStatus)

        this.keyIndex.Add count, NormaliseKey(varName)

NextVar:
    Next idx

    this.entryCount = count
    If count > 0 And count < maxEntries Then
        ReDim Preserve this.fieldKeys(1 To count)
        ReDim Preserve this.headers(1 To count)
        ReDim Preserve this.hidden(1 To count)
        ReDim Preserve this.mandatory(1 To count)
        ReDim Preserve this.locked(1 To count)
        ReDim Preserve this.verticalHeader(1 To count)
        ReDim Preserve this.positions(1 To count)
        ReDim Preserve this.controls(1 To count)
    End If
End Sub


'@section Helpers
'===============================================================================
'@description Private workers of the build and the accessors.

'@label:effective-hidden
'@sub-title Work out the state one entry really has
'@param index Long. A one based entry index.
'@return Boolean. True when the entry is hidden.
Private Function EffectiveHidden(ByVal index As Long) As Boolean
    If this.mandatory(index) Then
        EffectiveHidden = False
    ElseIf this.locked(index) Then
        EffectiveHidden = True
    Else
        EffectiveHidden = this.hidden(index)
    End If
End Function

'@label:matches-sheet
'@sub-title Whether a variable belongs to the sheet being built
'@param vars LLVariables. The variable reader.
'@param varName String. The variable to test.
'@return Boolean. True when the sheet names agree.
Private Function MatchesSheet(ByVal vars As LLVariables, ByVal varName As String) As Boolean
    Dim varSheetName As String

    varSheetName = Normalise(SafeValue(vars, COL_SHEET_NAME, varName))
    MatchesSheet = (StrComp(varSheetName, Normalise(this.sheetName), vbTextCompare) = 0)
End Function

'@label:matches-layer
'@sub-title Whether a dictionary sheet type suits the layer being built
'@details
'The dictionary vocabulary is "vlist1D" and "hlist2D". A printed sheet and a CRF
'are both derived from an hlist2D sheet, so all three of HList, Printed and CRF
'test for "hlist". The CRF case used to test for "crf", a word the vocabulary
'has never held, so every CRF list came back empty.
'@param sheetType String. The dictionary `sheet type` value, lower cased.
'@return Boolean. True when the variable belongs on this layer.
Private Function MatchesLayer(ByVal sheetType As String) As Boolean
    Select Case this.targetLayer
        Case ShowHideLayerHList, ShowHideLayerPrinted, ShowHideLayerCRF
            MatchesLayer = (InStr(1, sheetType, "hlist", vbBinaryCompare) > 0)
        Case ShowHideLayerVList
            MatchesLayer = (InStr(1, sheetType, "vlist", vbBinaryCompare) > 0)
        Case Else
            MatchesLayer = False
    End Select
End Function

'@label:status-for-layer
'@sub-title Read the status column that belongs to the layer
'@param vars LLVariables. The variable reader.
'@param varName String. The variable to read.
'@param registerBook String. The `register book` value, already read.
'@param hasVisibility Boolean. Whether the dictionary carries `visibility`.
'@param hasStatus Boolean. Whether the dictionary carries `status`.
'@param hasCrfStatus Boolean. Whether the dictionary carries `crf status`.
'@return String. The status text for this layer.
Private Function StatusForLayer(ByVal vars As LLVariables, _
                             ByVal varName As String, _
                             ByVal registerBook As String, _
                             ByVal hasVisibility As Boolean, _
                             ByVal hasStatus As Boolean, _
                             ByVal hasCrfStatus As Boolean) As String
    Select Case this.targetLayer
        Case ShowHideLayerPrinted
            StatusForLayer = registerBook
        Case ShowHideLayerCRF
            If hasCrfStatus Then StatusForLayer = SafeValue(vars, COL_CRF_STATUS, varName)
        Case Else
            If hasVisibility Then
                StatusForLayer = SafeValue(vars, COL_VISIBILITY, varName)
            ElseIf hasStatus Then
                StatusForLayer = SafeValue(vars, COL_STATUS, varName)
            End If
    End Select
End Function

'@label:is-mandatory-status
'@sub-title Whether a status makes an entry always visible
'@details
'The first geo level of an HList sheet is mandatory whatever its status says,
'because the geo cascade of every other level reads it.
'@param layerStatus String. The status of this layer, lower cased.
'@param controlVal String. The control value, lower cased.
'@param layer Byte. A ShowHideWorksheetLayer value. The layer being built.
'@return Boolean. True when the user cannot hide the entry.
Private Function IsMandatoryStatus(ByVal layerStatus As String, _
                                   ByVal controlVal As String, _
                                   ByVal layer As Byte) As Boolean
    If layer = ShowHideLayerHList And controlVal = "geo1" Then
        IsMandatoryStatus = True
        Exit Function
    End If

    Select Case layerStatus
        Case "always visible", "mandatory"
            IsMandatoryStatus = True
        Case Else
            IsMandatoryStatus = False
    End Select
End Function

'@label:is-locked-status
'@sub-title Whether a status holds an entry hidden
'@details
'A calculated column and an automatic list have nothing to give on a printed
'sheet or a CRF, so both layers hold them hidden. A mandatory entry is never
'locked as well.
'@param layerStatus String. The status of this layer, lower cased.
'@param controlVal String. The control value, lower cased.
'@param layer Byte. A ShowHideWorksheetLayer value. The layer being built.
'@param isMandatoryEntry Boolean. Whether the entry is already mandatory.
'@param hasListAuto Boolean. Whether the dictionary carries `list auto`.
'@param vars LLVariables. The variable reader.
'@param varName String. The variable being read.
'@return Boolean. True when the entry is held hidden.
Private Function IsLockedStatus(ByVal layerStatus As String, _
                                ByVal controlVal As String, _
                                ByVal layer As Byte, _
                                ByVal isMandatoryEntry As Boolean, _
                                ByVal hasListAuto As Boolean, _
                                ByVal vars As LLVariables, _
                                ByVal varName As String) As Boolean
    If isMandatoryEntry Then Exit Function

    If layerStatus = "always hidden" Then
        IsLockedStatus = True
        Exit Function
    End If

    If layer <> ShowHideLayerCRF And layer <> ShowHideLayerPrinted Then Exit Function

    Select Case controlVal
        Case "formula", "choice_formula", "case_when"
            IsLockedStatus = True
            Exit Function
    End Select

    If hasListAuto Then
        IsLockedStatus = (Normalise(SafeValue(vars, COL_LIST_AUTO, varName)) = "yes")
    End If
End Function

'@label:initial-hidden
'@sub-title The state a free entry starts in
'@details
'The `register book` vocabulary of a printed sheet is "print, horizontal
'header", "print, vertical header" and "hidden". An empty cell means the column
'prints, which is why the default answer is visible.
'@param layerStatus String. The status of this layer, lower cased.
'@return Boolean. True when the entry starts hidden.
Private Function InitialHidden(ByVal layerStatus As String) As Boolean
    Select Case layerStatus
        Case "hidden", "hidden on creation", "always hidden", _
             "optional, hidden", "do not print"
            InitialHidden = True
        Case Else
            InitialHidden = False
    End Select
End Function

'@label:safe-value
'@sub-title Read a dictionary cell and answer empty when it is absent
'@param vars LLVariables. The variable reader.
'@param colName String. The column to read.
'@param varName String. The variable to read.
'@return String. The cell value, or an empty string.
Private Function SafeValue(ByVal vars As LLVariables, _
                           ByVal colName As String, _
                           ByVal varName As String) As String
    Dim result As String

    On Error Resume Next
        result = vars.Value(colName, varName)
    On Error GoTo 0

    SafeValue = result
End Function

'@label:normalise
'@sub-title Trim and lower case a value for comparison
'@param value String. The text to normalise.
'@return String. The trimmed, lower cased text.
Private Function Normalise(ByVal value As String) As String
    Normalise = LCase$(Trim$(value))
End Function

'@label:normalise-key
'@sub-title Build the lookup key of a variable name
'@details
'The same shape on the way in and on the way out. A dictionary name carrying a
'trailing space used to be stored untrimmed and looked up trimmed, so its saved
'state was dropped on every load.
'@param value String. The variable name.
'@return String. The lookup key.
Private Function NormaliseKey(ByVal value As String) As String
    NormaliseKey = LCase$(Trim$(value))
End Function

'@label:ensure-valid-index
'@sub-title Refuse an entry index outside the list
'@param index Long. The index to check.
'@throws ProjectError.InvalidArgument When the index is out of range.
Private Sub EnsureValidIndex(ByVal index As Long)
    If index < 1 Or index > this.entryCount Then
        ThrowError ProjectError.InvalidArgument, _
                   "Entry index " & CStr(index) & " is outside range 1.." & CStr(this.entryCount)
    End If
End Sub

'@label:validate-layer
'@sub-title Refuse a layer outside the four values
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to check.
'@throws ProjectError.InvalidArgument When the layer is unknown.
Private Sub ValidateLayer(ByVal layer As Byte)
    If layer < ShowHideLayerHList Or layer > ShowHideLayerCRF Then
        ThrowError ProjectError.InvalidArgument, _
                   "There is no show/hide layer '" & CStr(layer) & "' in a linelist"
    End If
End Sub

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
