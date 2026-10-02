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

'@description
'What a span of positions - a section, say - looks like once every entry inside
'it has been read. Fixed and Shown are told apart on purpose: a span holding
'nothing the user owns cannot be toggled at all, and a caller that read it as
'Shown would offer an action that does nothing.
Public Enum ShowHideRangeState
    ShowHideRangeEmpty = 0
    ShowHideRangeShown = 1
    ShowHideRangeHidden = 2
    ShowHideRangeMixed = 3
    ShowHideRangeFixed = 4
End Enum


'@class ShowHide
'@description
'Holds the list of variables one worksheet layer offers the user, and the
'visible or hidden choice for each of them. The list is built from the
'dictionary. EntryCount, FieldKey, HeaderText, IsHidden, PositionIndex,
'ControlType and AuthoredVertical read one entry, IndexOf and HasField find
'one, SetHidden and SetAllOptionalHidden change the choices, SetHiddenInRange
'and RangeState do the same for a whole span of positions, Apply writes them
'through a layout, Adopt reads them back from a saved set and ResetToAuthored
'puts the choices, the sheet and the printed header directions back where the
'dictionary started them. The class reads the dictionary alone: the worksheet
'is ShowHideLayout's work and the saved choices are ShowHideStore's.
'
'A SECTION IS A SPAN OF POSITIONS
'-------------------------------------------------------------------------------
'Hiding a whole section is SetHiddenInRange over the positions SectionMap
'recorded for it. Nothing about a section is stored here or anywhere else: the
'choice lands on the member entries, so it travels through the store and
'through a migration file with no column of its own, and there is no second
'record that could disagree with the first.
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
'             column on a printed sheet or on a CRF lands here, and so does an
'             automatic list: both are named by their CONTROL, never by the
'             "list auto" column, which is about something else entirely.
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
'THE DICTIONARY IS READ IN BLOCKS
'-------------------------------------------------------------------------------
'The build reads each column it needs as one whole-column block through
'dict.DataRange and answers every per-variable question from those arrays.
'Every block spans the same rows (the data body of the dictionary), so one row
'index reads across all of them. It used to ask LLVariables.Value one cell at a
'time, two crossings per call, 2V + 14N per build; twelve block reads cost the
'same whatever V is.
'
'ONLY WHAT DIFFERS IS WRITTEN
'-------------------------------------------------------------------------------
'Beside hidden() the list keeps written(), the sheet state as Apply last wrote
'it or Adopt last read it, and stored(), the flags as the store last loaded or
'saved them. Apply writes the positions where the wanted state differs from
'written() and reads a position it has never seen before writing it, so a
'section press writes that section's columns alone. HasChanges compares
'hidden() with stored(), so a close where nothing moved can skip its save.
'@depends LLdictionary, ShowHideLayout, BetterArray, Checking
'@version 1.4 (2026-09-01)


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

'The one `register book` value that asks for a turned header
Private Const PRINT_VERTICAL As String = "print, vertical header"

'The dictionary columns the build reads, one block each, row-aligned. A column
'the dictionary does not carry is a block of empty strings and its flag is
'False, so the build reads the same shape whatever the dictionary holds.
Private Type TDictionaryBlocks
    rowCount As Long
    names As Variant
    sheetNames As Variant
    sheetTypes As Variant
    visibilities As Variant
    statuses As Variant
    controls As Variant
    labels As Variant
    subSections As Variant
    columnIndexes As Variant
    crfIndexes As Variant
    crfStatuses As Variant
    registerBooks As Variant
    hasVisibility As Boolean
    hasStatus As Boolean
    hasRegisterBook As Boolean
    hasCrfStatus As Boolean
    hasSubSection As Boolean
    hasSheetType As Boolean
    hasCrfIndex As Boolean
    hasColumnIndex As Boolean
End Type

'What the list knows about one position on the sheet: nothing yet, or the last
'state it wrote there or read from it.
Private Const WRITTEN_UNKNOWN As Byte = 0
Private Const WRITTEN_SHOWN As Byte = 1
Private Const WRITTEN_HIDDEN As Byte = 2

Private Type TShowHide
    fieldKeys() As String
    headers() As String
    hidden() As Boolean
    authoredHidden() As Boolean
    written() As Byte       'the sheet state as last written by Apply or read by Adopt
    stored() As Boolean     'the hidden flag as the store last loaded or saved it
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

'@label:set-hidden-in-range
'@sub-title Hide or show every entry sitting in a span of positions
'@details
'What hiding a whole section comes down to. The span is a range of column
'numbers on an HList sheet and of row numbers on a VList one, which is what
'SectionMap records for each section.
'
'Hiding a section stores nothing of its own: the choice lands on each member
'entry, so the store, the show/hide form and a migration file all carry a
'collapsed section without knowing the word "section". Nothing can drift out of
'step with anything else, because there is only one record of the choice.
'
'A mandatory or locked entry inside the span keeps its state, as it does
'everywhere else. So a section holding a mandatory variable never disappears
'whole, and the count says how many entries actually moved.
'@param firstPos Long. One end of the span.
'@param lastPos Long. The other end of the span.
'@param hidden Boolean. True to hide the span.
'@return Long. How many entries the call changed.
'@export
Public Function SetHiddenInRange(ByVal firstPos As Long, _
                                 ByVal lastPos As Long, _
                                 ByVal hidden As Boolean) As Long
    Dim idx As Long
    Dim changed As Long

    For idx = 1 To this.entryCount
        If PositionInRange(this.positions(idx), firstPos, lastPos) Then
            If Not (this.mandatory(idx) Or this.locked(idx)) Then
                If this.hidden(idx) <> hidden Then
                    this.hidden(idx) = hidden
                    changed = changed + 1
                End If
            End If
        End If
    Next idx

    SetHiddenInRange = changed
End Function

'@label:range-state
'@sub-title Where a span of positions stands
'@details
'Read from the entries the user owns, so a section whose only visible column is
'a mandatory one still answers Hidden and a caller toggling it offers Show.
'@param firstPos Long. One end of the span.
'@param lastPos Long. The other end of the span.
'@return Byte. A ShowHideRangeState value. Empty when no entry sits in the span,
'  Fixed when none of those that do is free, and otherwise Shown, Hidden or
'  Mixed by what the free entries say.
'@export
Public Function RangeState(ByVal firstPos As Long, ByVal lastPos As Long) As Byte
    Dim idx As Long
    Dim inRange As Long
    Dim freeCount As Long
    Dim hiddenCount As Long

    For idx = 1 To this.entryCount
        If PositionInRange(this.positions(idx), firstPos, lastPos) Then
            inRange = inRange + 1
            If Not (this.mandatory(idx) Or this.locked(idx)) Then
                freeCount = freeCount + 1
                If this.hidden(idx) Then hiddenCount = hiddenCount + 1
            End If
        End If
    Next idx

    If inRange = 0 Then
        RangeState = ShowHideRangeEmpty
    ElseIf freeCount = 0 Then
        RangeState = ShowHideRangeFixed
    ElseIf hiddenCount = 0 Then
        RangeState = ShowHideRangeShown
    ElseIf hiddenCount = freeCount Then
        RangeState = ShowHideRangeHidden
    Else
        RangeState = ShowHideRangeMixed
    End If
End Function


'@section Worksheet
'===============================================================================
'@description The two calls that put the list and a sheet in step.

'@label:apply
'@sub-title Put the worksheet in the state the list describes
'@details
'The one place the entry list reaches a sheet, and it goes through the layout,
'so the row versus column question and the protection bracket are answered in
'one place. Entries with no position are skipped.
'
'Only the positions that differ are written. The list remembers what it last
'wrote to each position and what Adopt last read there; a position it knows
'nothing about yet is read off the sheet first, which costs a read where a
'write used to go. So a section press writes the columns of that section
'alone, and an Apply on a sheet already in step writes nothing. The force
'flag writes every position whatever the list remembers, for a caller landing
'on a sheet whose state it cannot trust, such as the migration import.
'@param layout ShowHideLayout. The sheet to reconcile.
'@param force Optional Boolean. True writes every position. Default False.
'@return Long. The number of positions that were written.
'@export
Public Function Apply(ByVal layout As ShowHideLayout, _
                      Optional ByVal force As Boolean = False) As Long
    Dim idx As Long
    Dim position As Long
    Dim wanted As Boolean
    Dim applied As Long

    If layout Is Nothing Then Exit Function

    layout.BeginBatch
    For idx = 1 To this.entryCount
        position = this.positions(idx)
        If position > 0 Then
            wanted = EffectiveHidden(idx)

            If Not force And this.written(idx) = WRITTEN_UNKNOWN Then
                RememberWritten idx, layout.IsHidden(position)
            End If

            If force Or WrittenHidden(idx) <> wanted Then
                layout.SetHidden position, wanted
                RememberWritten idx, wanted
                applied = applied + 1
            End If
        End If
    Next idx
    layout.EndBatch

    Apply = applied
End Function

'@label:adopt
'@sub-title Read the worksheet back into the list
'@details
'For the case where the user hid a column by hand rather than through the form.
'Mandatory and locked entries keep what the dictionary said. Every position is
'read all the same, so the list knows the whole sheet and the next Apply
'writes only what differs from it.
'@param layout ShowHideLayout. The sheet to read.
'@return Long. The number of free entries that were read.
'@export
Public Function Adopt(ByVal layout As ShowHideLayout) As Long
    Dim idx As Long
    Dim position As Long
    Dim onSheet As Boolean
    Dim adopted As Long

    If layout Is Nothing Then Exit Function

    For idx = 1 To this.entryCount
        position = this.positions(idx)
        If position > 0 Then
            onSheet = layout.IsHidden(position)
            RememberWritten idx, onSheet
            If Not (this.mandatory(idx) Or this.locked(idx)) Then
                this.hidden(idx) = onSheet
                adopted = adopted + 1
            End If
        End If
    Next idx

    Adopt = adopted
End Function

'@label:reset-to-authored
'@sub-title Put the choices and the sheet back where the dictionary started them
'@details
'The reset button of the advanced form comes through here, once per worksheet.
'Every free entry goes back to the hidden or visible state the dictionary
'authored, the sheet is put in step through Apply, and on a printed sheet each
'header direction goes back to what `register book` asked. Mandatory and locked
'entries never left their authored state, so they are covered by the same walk.
'@param layout ShowHideLayout. The sheet to reconcile, or Nothing to reset the
'  choices alone.
'@return Long. The number of positions Apply set.
'@export
Public Function ResetToAuthored(ByVal layout As ShowHideLayout) As Long
    Dim idx As Long
    Dim position As Long

    For idx = 1 To this.entryCount
        this.hidden(idx) = this.authoredHidden(idx)
    Next idx

    ResetToAuthored = Apply(layout)

    If layout Is Nothing Then Exit Function
    If Not layout.SupportsOrientation Then Exit Function

    layout.BeginBatch
    For idx = 1 To this.entryCount
        position = this.positions(idx)
        If position > 0 Then
            layout.SetOrientation position, this.verticalHeader(idx)
        End If
    Next idx
    layout.EndBatch
End Function


'@section Store
'===============================================================================
'@description What the store last saw, so a close with nothing changed saves nothing.

'@label:has-changes
'@prop-title Whether any choice differs from what the store last loaded or saved
'@details
'Read before a save. The baseline is set by the build (the authored state,
'which is what an empty store means), and moved by ShowHideStore.Load and
'ShowHideStore.Save through MarkStored. Adopt leaves it alone on purpose: a
'column the user hid by hand is a change the store has not seen yet, and the
'next save is where it gets recorded.
'@export
Public Property Get HasChanges() As Boolean
    Dim idx As Long

    For idx = 1 To this.entryCount
        If this.hidden(idx) <> this.stored(idx) Then
            HasChanges = True
            Exit Property
        End If
    Next idx
End Property

'@label:mark-stored
'@sub-title Take the choices as they stand for what the store holds
'@details
'ShowHideStore calls it at the end of a Load and of a Save of the current
'state, so HasChanges answers False until the next choice moves.
'@export
Public Sub MarkStored()
    Dim idx As Long

    For idx = 1 To this.entryCount
        this.stored(idx) = this.hidden(idx)
    Next idx
End Sub


'@section Build
'===============================================================================
'@description The dictionary walk the factory runs.

'@label:initialise
'@sub-title Fill the entry list from the dictionary
'@details
'Public because the factory calls it on a fresh instance. The columns the
'build needs are read as whole blocks first, then every row is read from the
'arrays, and the entry arrays are trimmed to what matched.
'@param dict LLdictionary. The dictionary to read.
'@param layer Byte. A ShowHideWorksheetLayer value. The layer being built.
'@param sheetName String. The base sheet name.
'@export
Public Sub Initialise(ByVal dict As LLdictionary, _
                      ByVal layer As Byte, _
                      ByVal sheetName As String)

    Dim blocks As TDictionaryBlocks
    Dim idx As Long
    Dim varName As String
    Dim count As Long
    Dim maxEntries As Long
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

    ReadBlocks dict, blocks
    If blocks.rowCount = 0 Then Exit Sub

    maxEntries = blocks.rowCount
    ReDim this.fieldKeys(1 To maxEntries)
    ReDim this.headers(1 To maxEntries)
    ReDim this.hidden(1 To maxEntries)
    ReDim this.authoredHidden(1 To maxEntries)
    ReDim this.mandatory(1 To maxEntries)
    ReDim this.locked(1 To maxEntries)
    ReDim this.verticalHeader(1 To maxEntries)
    ReDim this.positions(1 To maxEntries)
    ReDim this.controls(1 To maxEntries)
    ReDim this.written(1 To maxEntries)
    ReDim this.stored(1 To maxEntries)

    count = 0

    For idx = 1 To blocks.rowCount
        varName = BlockText(blocks.names, idx)
        If LenB(Trim$(varName)) = 0 Then GoTo NextVar

        If Not MatchesSheet(BlockText(blocks.sheetNames, idx)) Then GoTo NextVar

        If blocks.hasSheetType Then
            If Not MatchesLayer(Normalise(BlockText(blocks.sheetTypes, idx))) Then GoTo NextVar
        End If

        'A variable the designer hid never reaches the user's list
        If blocks.hasVisibility Then
            If Normalise(BlockText(blocks.visibilities, idx)) = "always hidden" Then GoTo NextVar
        End If

        If HasField(varName) Then GoTo NextVar

        count = count + 1

        this.fieldKeys(count) = varName

        controlVal = Normalise(BlockText(blocks.controls, idx))
        this.controls(count) = controlVal

        headerVal = BlockText(blocks.labels, idx)
        If LenB(headerVal) = 0 Then headerVal = varName

        If blocks.hasSubSection Then
            If InStr(1, controlVal, "geo") > 0 Or InStr(1, controlVal, "hf") > 0 Then
                subSectionVal = BlockText(blocks.subSections, idx)
                If LenB(subSectionVal) > 0 Then
                    headerVal = subSectionVal & ": " & headerVal
                End If
            End If
        End If
        this.headers(count) = headerVal

        posValue = vbNullString
        If layer = ShowHideLayerCRF And blocks.hasCrfIndex Then
            posValue = BlockText(blocks.crfIndexes, idx)
        ElseIf blocks.hasColumnIndex Then
            posValue = BlockText(blocks.columnIndexes, idx)
        End If
        If LenB(posValue) > 0 And IsNumeric(posValue) Then
            this.positions(count) = CLng(posValue)
        Else
            this.positions(count) = 0
        End If

        registerBook = vbNullString
        If blocks.hasRegisterBook Then registerBook = Normalise(BlockText(blocks.registerBooks, idx))
        this.verticalHeader(count) = (registerBook = PRINT_VERTICAL)

        layerStatus = Normalise(StatusForLayer(blocks, idx, registerBook))

        this.mandatory(count) = IsMandatoryStatus(layerStatus, controlVal, layer)
        this.locked(count) = IsLockedStatus(layerStatus, controlVal, layer, _
                                            this.mandatory(count))
        this.hidden(count) = InitialHidden(layerStatus)
        this.authoredHidden(count) = this.hidden(count)
        this.written(count) = WRITTEN_UNKNOWN
        this.stored(count) = this.hidden(count)

        this.keyIndex.Add count, NormaliseKey(varName)

NextVar:
    Next idx

    this.entryCount = count
    If count > 0 And count < maxEntries Then
        ReDim Preserve this.fieldKeys(1 To count)
        ReDim Preserve this.headers(1 To count)
        ReDim Preserve this.hidden(1 To count)
        ReDim Preserve this.authoredHidden(1 To count)
        ReDim Preserve this.mandatory(1 To count)
        ReDim Preserve this.locked(1 To count)
        ReDim Preserve this.verticalHeader(1 To count)
        ReDim Preserve this.positions(1 To count)
        ReDim Preserve this.controls(1 To count)
        ReDim Preserve this.written(1 To count)
        ReDim Preserve this.stored(1 To count)
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

'@label:remember-written
'@sub-title Record what the sheet holds at one entry's position
'@param index Long. A one based entry index.
'@param hidden Boolean. The state written to or read from the sheet.
Private Sub RememberWritten(ByVal index As Long, ByVal hidden As Boolean)
    If hidden Then
        this.written(index) = WRITTEN_HIDDEN
    Else
        this.written(index) = WRITTEN_SHOWN
    End If
End Sub

'@label:written-hidden
'@sub-title What the list last saw at one entry's position
'@param index Long. A one based entry index.
'@return Boolean. True when the last write or read found it hidden.
Private Function WrittenHidden(ByVal index As Long) As Boolean
    WrittenHidden = (this.written(index) = WRITTEN_HIDDEN)
End Function

'@label:position-in-range
'@sub-title Whether one entry sits inside a span of positions
'@details
'The two bounds are read in either order, so a caller that hands them over the
'other way round gets the same answer. An entry with no position, which is what
'a variable the dictionary gives no `column index` looks like, is in no span.
'@param position Long. The entry position.
'@param firstPos Long. One end of the span.
'@param lastPos Long. The other end of the span.
'@return Boolean. True when the position falls inside the span.
Private Function PositionInRange(ByVal position As Long, _
                                 ByVal firstPos As Long, _
                                 ByVal lastPos As Long) As Boolean
    If position < 1 Then Exit Function

    If firstPos <= lastPos Then
        PositionInRange = (position >= firstPos) And (position <= lastPos)
    Else
        PositionInRange = (position >= lastPos) And (position <= firstPos)
    End If
End Function

'@label:matches-sheet
'@sub-title Whether a variable belongs to the sheet being built
'@param varSheetName String. The variable's `sheet name` value.
'@return Boolean. True when the sheet names agree.
Private Function MatchesSheet(ByVal varSheetName As String) As Boolean
    MatchesSheet = (StrComp(Normalise(varSheetName), Normalise(this.sheetName), vbTextCompare) = 0)
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
'@param blocks TDictionaryBlocks. The column blocks of the dictionary.
'@param row Long. The dictionary row of the variable, one based.
'@param registerBook String. The `register book` value, already read.
'@return String. The status text for this layer.
Private Function StatusForLayer(ByRef blocks As TDictionaryBlocks, _
                                ByVal row As Long, _
                                ByVal registerBook As String) As String
    Select Case this.targetLayer
        Case ShowHideLayerPrinted
            StatusForLayer = registerBook
        Case ShowHideLayerCRF
            If blocks.hasCrfStatus Then StatusForLayer = BlockText(blocks.crfStatuses, row)
        Case Else
            If blocks.hasVisibility Then
                StatusForLayer = BlockText(blocks.visibilities, row)
            ElseIf blocks.hasStatus Then
                StatusForLayer = BlockText(blocks.statuses, row)
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
'@return Boolean. True when the entry is held hidden.
Private Function IsLockedStatus(ByVal layerStatus As String, _
                                ByVal controlVal As String, _
                                ByVal layer As Byte, _
                                ByVal isMandatoryEntry As Boolean) As Boolean
    If isMandatoryEntry Then Exit Function

    If layerStatus = "always hidden" Then
        IsLockedStatus = True
        Exit Function
    End If

    If layer <> ShowHideLayerCRF And layer <> ShowHideLayerPrinted Then Exit Function

    'An automatic list is named by its CONTROL, the same way the calculated
    'controls beside it are. The test used to read the "list auto" DICTIONARY
    'COLUMN instead, which is a different thing entirely: that column holds
    '"list_auto_origin", "list_auto_indirect_origin" or nothing, and says which
    'variables an automatic list is collected FROM. It was compared against
    '"yes", a value that column has never held, so no automatic list was ever
    'held hidden here -- and had it matched, it would have hidden the origins
    'rather than the lists.
    Select Case controlVal
        Case "formula", "choice_formula", "case_when", "list_auto"
            IsLockedStatus = True
            Exit Function
    End Select
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

'@label:read-blocks
'@sub-title Read every column the build needs as one block each
'@details
'The row count comes from the `variable name` column, and every other block is
'read over the same rows, because dict.DataRange spans the data body of the
'dictionary whichever column is asked for. A column the dictionary lacks is
'read as a block of empty strings with its flag False. `variable name` is a
'required column, so a dictionary that has no such column has no rows to give.
'@param dict LLdictionary. The dictionary to read.
'@param blocks TDictionaryBlocks. Filled here.
Private Sub ReadBlocks(ByVal dict As LLdictionary, ByRef blocks As TDictionaryBlocks)
    blocks.rowCount = 0
    blocks.names = ColumnBlock(dict, "variable name", 0, blocks.rowCount)
    If blocks.rowCount = 0 Then Exit Sub

    blocks.hasVisibility = dict.ColumnExists(COL_VISIBILITY)
    blocks.hasStatus = dict.ColumnExists(COL_STATUS)
    blocks.hasRegisterBook = dict.ColumnExists(COL_REGISTER_BOOK)
    blocks.hasCrfStatus = dict.ColumnExists(COL_CRF_STATUS)
    blocks.hasSubSection = dict.ColumnExists(COL_SUB_SECTION)
    blocks.hasSheetType = dict.ColumnExists(COL_SHEET_TYPE)
    blocks.hasCrfIndex = dict.ColumnExists(COL_CRF_INDEX)
    blocks.hasColumnIndex = dict.ColumnExists(COL_COLUMN_INDEX)

    blocks.sheetNames = ColumnBlock(dict, COL_SHEET_NAME, blocks.rowCount)
    blocks.controls = ColumnBlock(dict, COL_CONTROL, blocks.rowCount)
    blocks.labels = ColumnBlock(dict, COL_MAIN_LABEL, blocks.rowCount)
    blocks.sheetTypes = ColumnBlock(dict, COL_SHEET_TYPE, blocks.rowCount, , blocks.hasSheetType)
    blocks.visibilities = ColumnBlock(dict, COL_VISIBILITY, blocks.rowCount, , blocks.hasVisibility)
    blocks.statuses = ColumnBlock(dict, COL_STATUS, blocks.rowCount, , blocks.hasStatus)
    blocks.subSections = ColumnBlock(dict, COL_SUB_SECTION, blocks.rowCount, , blocks.hasSubSection)
    blocks.columnIndexes = ColumnBlock(dict, COL_COLUMN_INDEX, blocks.rowCount, , blocks.hasColumnIndex)
    blocks.crfIndexes = ColumnBlock(dict, COL_CRF_INDEX, blocks.rowCount, , blocks.hasCrfIndex)
    blocks.crfStatuses = ColumnBlock(dict, COL_CRF_STATUS, blocks.rowCount, , blocks.hasCrfStatus)
    blocks.registerBooks = ColumnBlock(dict, COL_REGISTER_BOOK, blocks.rowCount, , blocks.hasRegisterBook)
End Sub

'@label:column-block
'@sub-title One dictionary column as a one based array of strings
'@details
'One crossing per column: the range's Value comes back as a 2D array, or as a
'single value when the dictionary holds one row, and both are flattened to the
'same shape. A column that is absent, or that the read refuses, answers a
'block of `expectedRows` empty strings, so a caller reads it like any other.
'@param dict LLdictionary. The dictionary to read.
'@param colName String. The column to read.
'@param expectedRows Long. The rows every block has to hold; 0 lets the column
'  decide, which is how the first block sets the count.
'@param rowsRead Long. Set to the number of rows the block holds.
'@param present Optional Boolean. False skips the read and answers the empty block.
'@return Variant. A one based array of strings.
Private Function ColumnBlock(ByVal dict As LLdictionary, _
                             ByVal colName As String, _
                             ByVal expectedRows As Long, _
                             Optional ByRef rowsRead As Long, _
                             Optional ByVal present As Boolean = True) As Variant
    Dim source As Range
    Dim raw As Variant
    Dim block() As String
    Dim nbRows As Long
    Dim counter As Long

    If present Then
        On Error Resume Next
            Set source = dict.DataRange(colName)
        On Error GoTo 0
    End If

    If source Is Nothing Then
        nbRows = expectedRows
    Else
        nbRows = source.Rows.Count
        If expectedRows > 0 Then nbRows = expectedRows
    End If

    rowsRead = nbRows
    If nbRows < 1 Then
        ColumnBlock = Array()
        Exit Function
    End If

    ReDim block(1 To nbRows)
    If source Is Nothing Then
        ColumnBlock = block
        Exit Function
    End If

    If source.Rows.Count = 1 Then
        block(1) = CellText(source.Cells(1, 1).Value)
    Else
        raw = source.Columns(1).Value
        For counter = 1 To nbRows
            If counter <= UBound(raw, 1) Then block(counter) = CellText(raw(counter, 1))
        Next counter
    End If

    ColumnBlock = block
End Function

'@label:block-text
'@sub-title One row of a block, empty when the block holds nothing
'@param block Variant. A block ColumnBlock built.
'@param row Long. The row wanted, one based.
'@return String. The text at that row.
Private Function BlockText(ByRef block As Variant, ByVal row As Long) As String
    If Not IsArray(block) Then Exit Function
    On Error Resume Next
        BlockText = block(row)
    On Error GoTo 0
End Function

'@label:cell-text
'@sub-title A cell value as text, with an error value or an empty read as nothing
'@param value Variant. What the cell answered.
'@return String. The text.
Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    If IsNull(value) Then Exit Function
    CellText = CStr(value)
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
