VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The show/hide choices a user made last time"

'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("The show/hide choices a user made last time")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit

'@class ShowHideStore
'@description
'The table that carries a user's show/hide choices from one session to the
'next, and from one workbook to another during a migration. One table holds all
'four layers, told apart by the `layer` column, and every saved layout, told
'apart by the `layout_name` column. Save writes choices, Load reads them back,
'Clear drops them, HasLayer says whether a layer has rows, and HasTable, Table
'and RowCount answer where the store stands.
'
'THE STORE MAKES ITS OWN TABLE
'-------------------------------------------------------------------------------
'Create and CreateOnSheet build the ListObject when it is absent, and add the
'`layout_name` column to a table an older build made without it.
'
'CreateForRead binds to whatever table a workbook already carries and builds
'nothing, which is what reading a migration file needs.
'
'SEVEN COLUMNS
'-------------------------------------------------------------------------------
'  layer         hlist / printed / vlist / crf
'  field_key     the variable name, and the key rows are matched on
'  header_text   for a person reading the sheet
'  hidden_flag   true / false
'  entry_size    the width or the height the entry has when shown, blank when
'                unknown
'  orientation   vertical / horizontal, blank off the printed layer
'  layout_name   empty on the rows of the current state, the layout's name on
'                the rows of a saved layout
'
'The size and the hidden flag are two separate columns. A size at or below zero
'means "leave the sheet alone".
'
'A file written by an older build carries `column_width` where this one writes
'`entry_size`, and no `layout_name` at all. Load reads either size name, and
'every row of a file without the layout column reads as current state.
'
'SAVED LAYOUTS
'-------------------------------------------------------------------------------
'A saved layout is a named copy of the same rows. Save and Load take the name;
'LayoutNames, HasLayout and LayoutCount answer what is stored; DeleteLayout and
'RenameLayout manage the list; MergeLayoutsFrom copies another store's layouts
'in, which is how they travel through a migration file. The store holds at
'most MaxSavedLayouts names, and Save refuses the one past the cap.
'
'SAVE REPLACES BY KEY
'-------------------------------------------------------------------------------
'Save replaces the rows whose layer, layout name and field key match the entry
'list, and keeps every other row. Several worksheets share one layer - every
'HList sheet writes `hlist` rows - so replacing a whole layer would drop the
'other sheets' rows on every save. Variable names are unique across the
'workbook, which is what makes the key safe. A row whose variable left the
'dictionary lingers; Load skips it.
'@depends ShowHide, ShowHideLayout, Checking, BetterArray
'@version 1.2 (2026-09-01)


Private Const CLASS_NAME As String = "ShowHideStore"

'The worksheet a linelist keeps its show/hide table on, and the name a
'linelist generated before the internal-sheet rename gave it
Private Const STORE_SHEET As String = "__show_hide"
Private Const LEGACY_SHEET As String = "show_hide__"

'The table this class writes, and the name an older build wrote
Private Const STORE_TABLE As String = "show_hide_state"
Private Const LEGACY_TABLE As String = "show_hide_migration"

'The per-sheet tables the first show/hide design tiled across the same
'worksheet. Both designs answer to ListObjects(1), so they cannot share a sheet.
Private Const LEGACY_PREFIX As String = "ShowHideTable_"

Private Const COL_LAYER As String = "layer"
Private Const COL_FIELD_KEY As String = "field_key"
Private Const COL_HEADER As String = "header_text"
Private Const COL_HIDDEN As String = "hidden_flag"
Private Const COL_SIZE As String = "entry_size"
Private Const COL_ORIENTATION As String = "orientation"
Private Const COL_LAYOUT As String = "layout_name"
Private Const COL_LEGACY_SIZE As String = "column_width"

Private Const COLUMN_COUNT As Long = 7

'How many saved layouts one store holds at most
Private Const MAX_SAVED_LAYOUTS As Long = 10

Private Type TShowHideStore
    lo As ListObject
    sh As Worksheet
    droppedSizes As Collection   'the sizes of the rows a running Save replaces, keyed on the field
End Type

Private this As TShowHideStore


'@section Instantiation
'===============================================================================
'@description Three ways in: two that provision, one that reads what is there.

'@label:create
'@sub-title Open the store of a workbook, making the table when it is absent
'@param wkb Workbook. The workbook holding the __show_hide worksheet.
'@return ShowHideStore. A store ready to save and load.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@throws ProjectError.ElementNotFound When the workbook has no __show_hide
'  worksheet.
'@export
Public Function Create(ByVal wkb As Workbook) As ShowHideStore
    Dim sh As Worksheet

    If wkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A workbook is needed to find the show/hide store"
    End If

    'A workbook generated before the internal-sheet rename still carries the
    'old trailing name, so the new name is looked for first and the old one
    'answers when it is absent.
    On Error Resume Next
    Set sh = wkb.Worksheets(STORE_SHEET)
    If sh Is Nothing Then Set sh = wkb.Worksheets(LEGACY_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "This workbook has no " & STORE_SHEET & " worksheet"
    End If

    Set Create = CreateOnSheet(sh)
End Function

'@label:create-on-sheet
'@sub-title Open the store on one worksheet, making the table when it is absent
'@details
'Any per-sheet table of the first show/hide design is deleted on the way, since
'both designs read ListObjects(1) and the two layouts hold different columns.
'@param sh Worksheet. The sheet to keep the table on.
'@return ShowHideStore. A store ready to save and load.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@export
Public Function CreateOnSheet(ByVal sh As Worksheet) As ShowHideStore
    Dim instance As ShowHideStore

    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A worksheet is needed to keep the show/hide store on"
    End If

    Set instance = New ShowHideStore
    instance.Bind sh, Provision(sh)
    Set CreateOnSheet = instance
End Function

'@label:create-for-read
'@sub-title Bind to the table a worksheet already carries
'@details
'Provisions nothing, which is what reading another workbook's migration file
'needs. HasTable says whether anything was found.
'@param sh Worksheet. The sheet to read.
'@return ShowHideStore. A store bound to what is there.
'@export
Public Function CreateForRead(ByVal sh As Worksheet) As ShowHideStore
    Dim instance As ShowHideStore

    Set instance = New ShowHideStore
    If sh Is Nothing Then
        Set CreateForRead = instance
        Exit Function
    End If

    instance.Bind sh, FindTable(sh)
    Set CreateForRead = instance
End Function

'@label:bind
'@sub-title Take the sheet and the table the factory resolved
'@details
'Public because the factories call it on a fresh instance.
'@param sh Worksheet. The sheet the table lives on.
'@param lo ListObject. The table, or Nothing when the sheet carries none.
'@export
Public Sub Bind(ByVal sh As Worksheet, ByVal lo As ListObject)
    Set this.sh = sh
    Set this.lo = lo
End Sub


'@section Accessors
'===============================================================================
'@description What the store is bound to.

'@label:has-table
'@prop-title Whether a table was found or made
'@export
Public Property Get HasTable() As Boolean
    HasTable = Not (this.lo Is Nothing)
End Property

'@label:table
'@prop-title The ListObject the store reads and writes
'@export
Public Property Get Table() As ListObject
    Set Table = this.lo
End Property

'@label:row-count
'@prop-title How many rows the store holds, over all four layers
'@export
Public Property Get RowCount() As Long
    If this.lo Is Nothing Then Exit Property
    If this.lo.DataBodyRange Is Nothing Then Exit Property
    RowCount = this.lo.DataBodyRange.Rows.Count
End Property

'@label:has-layer
'@sub-title Whether the store holds any row of one layer
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to look for.
'@return Boolean. True when at least one row carries that layer.
'@export
Public Function HasLayer(ByVal layer As Byte) As Boolean
    Dim tableRows As Variant
    Dim layerCol As Long
    Dim token As String
    Dim counter As Long

    If this.lo Is Nothing Then Exit Function
    If this.lo.DataBodyRange Is Nothing Then Exit Function

    layerCol = ColumnIndex(COL_LAYER)
    If layerCol = 0 Then Exit Function

    token = LayerToken(layer)
    tableRows = this.lo.DataBodyRange.Value

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If StrComp(CStr(tableRows(counter, layerCol)), token, vbTextCompare) = 0 Then
            HasLayer = True
            Exit Function
        End If
    Next counter
End Function


'@section Saving
'===============================================================================
'@description Writing an entry list into the table.

'@label:save
'@sub-title Write one layer's choices, leaving every other row alone
'@details
'The rows whose layer, layout name and field key match the entry list are
'replaced; everything else is carried over untouched. So four layers share one
'table, several worksheets share one layer, and the current state sits beside
'the saved layouts.
'
'Without a layout the store writes visibility alone, which is what closing the
'form needs. With one it also records the size each entry has when shown and,
'on a printed sheet, the direction of its header - which is what a migration
'export and a saved layout need.
'@param entries ShowHide. The entry list to write.
'@param layout Optional ShowHideLayout. The sheet to read sizes from.
'@param layoutName Optional String. Empty writes the current state; a name
'  writes a saved layout.
'@throws ProjectError.ObjectNotInitialized When the entry list is Nothing.
'@throws ProjectError.ElementNotFound When the store has no table, or a name is
'  given and the table has no layout_name column.
'@throws ProjectError.InvalidArgument When a new name would pass the cap.
'@export
Public Sub Save(ByVal entries As ShowHide, _
                Optional ByVal layout As ShowHideLayout, _
                Optional ByVal layoutName As String = vbNullString)
    Dim kept As Variant
    Dim keptCount As Long
    Dim block As Variant
    Dim token As String
    Dim idx As Long
    Dim target As Long
    Dim position As Long
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim headerCol As Long
    Dim hiddenCol As Long
    Dim sizeCol As Long
    Dim orientCol As Long
    Dim layoutCol As Long
    Dim entrySize As Double

    If entries Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "There is no show/hide entry list to save"
    End If

    RequireTable

    layerCol = ColumnIndex(COL_LAYER)
    fieldCol = ColumnIndex(COL_FIELD_KEY)
    headerCol = ColumnIndex(COL_HEADER)
    hiddenCol = ColumnIndex(COL_HIDDEN)
    sizeCol = SizeColumnIndex()
    orientCol = ColumnIndex(COL_ORIENTATION)
    layoutCol = ColumnIndex(COL_LAYOUT)

    If layerCol = 0 Or fieldCol = 0 Or hiddenCol = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The show/hide table is missing one of layer, field_key or hidden_flag"
    End If

    layoutName = Trim$(layoutName)
    If LenB(layoutName) > 0 Then
        If layoutCol = 0 Then
            ThrowError ProjectError.ElementNotFound, _
                       "The show/hide table has no " & COL_LAYOUT & " column"
        End If
        If Not HasLayout(layoutName) Then
            If LayoutCount() >= MAX_SAVED_LAYOUTS Then
                ThrowError ProjectError.InvalidArgument, _
                           "The store already holds " & CStr(MAX_SAVED_LAYOUTS) & _
                           " saved layouts; delete one before saving another"
            End If
        End If
    End If

    token = LayerToken(entries.TargetLayer)
    'KeptRows files the size of every row it drops while this is set
    Set this.droppedSizes = New Collection
    kept = KeptRows(token, layerCol, layoutName, layoutCol, entries, fieldCol, keptCount)

    If keptCount + entries.EntryCount = 0 Then
        Set this.droppedSizes = Nothing
        WriteAll kept, 0
        Exit Sub
    End If

    ReDim block(1 To keptCount + entries.EntryCount, 1 To this.lo.ListColumns.Count)

    For idx = 1 To keptCount
        For target = 1 To UBound(block, 2)
            block(idx, target) = kept(idx, target)
        Next target
    Next idx

    If Not layout Is Nothing Then layout.BeginBatch

    For idx = 1 To entries.EntryCount
        target = keptCount + idx
        block(target, layerCol) = token
        block(target, fieldCol) = entries.FieldKey(idx)
        If headerCol > 0 Then block(target, headerCol) = entries.HeaderText(idx)
        block(target, hiddenCol) = BoolToString(entries.IsHidden(idx))
        If layoutCol > 0 And LenB(layoutName) > 0 Then block(target, layoutCol) = layoutName

        If Not layout Is Nothing Then
            position = entries.PositionIndex(idx)
            If position > 0 Then
                If sizeCol > 0 Then
                    entrySize = layout.SizeWhenShown(position)
                    'A layout that answers 0 for a hidden entry knows nothing of
                    'its size; the size the table already held for it stands.
                    If entrySize <= 0 Then entrySize = DroppedSize(entries.FieldKey(idx))
                    If entrySize > 0 Then block(target, sizeCol) = entrySize
                End If
                If orientCol > 0 And layout.SupportsOrientation Then
                    block(target, orientCol) = OrientationToken(layout.IsVertical(position))
                End If
            End If
        End If
    Next idx

    If Not layout Is Nothing Then layout.EndBatch

    WriteAll block, keptCount + entries.EntryCount
    Set this.droppedSizes = Nothing

    'The current state now matches the list; a named layout is a copy beside it
    If LenB(layoutName) = 0 Then entries.MarkStored
End Sub

'@label:clear
'@sub-title Drop every row of one layer of one state
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to drop.
'@param layoutName Optional String. Empty drops current-state rows; a name
'  drops that layout's rows of the layer. DeleteLayout drops a whole layout.
'@throws ProjectError.ElementNotFound When the store has no table.
'@export
Public Sub Clear(ByVal layer As Byte, _
                 Optional ByVal layoutName As String = vbNullString)
    Dim kept As Variant
    Dim keptCount As Long
    Dim layerCol As Long

    RequireTable

    layerCol = ColumnIndex(COL_LAYER)
    If layerCol = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The show/hide table is missing its " & COL_LAYER & " column"
    End If

    kept = KeptRows(LayerToken(layer), layerCol, Trim$(layoutName), _
                    ColumnIndex(COL_LAYOUT), Nothing, 0, keptCount)

    WriteAll kept, keptCount
End Sub


'@section Loading
'===============================================================================
'@description Reading the table back into an entry list.

'@label:load
'@sub-title Read one layer's saved choices into an entry list
'@details
'Only the rows of the entry list's own layer and of the named state are read.
'A variable the list does not carry is skipped, which is what happens when the
'dictionary changed between two sessions. On a file written before saved
'layouts existed, every row reads as current state.
'
'With a layout, the saved size and header direction are applied to the sheet as
'well. Visibility is left to ShowHide.Apply, so a caller that wants the sheet to
'match the list calls Load and then Apply.
'
'The sizes are written only when writeSizes says so. The open of the form and a
'section press pass False: the live sheet already carries the sizes the store
'recorded at the last close, and N size writes would put back what is there. A
'layout restore and a migration import pass True, because their sizes come from
'elsewhere.
'@param entries ShowHide. The entry list to fill.
'@param layout Optional ShowHideLayout. The sheet to size.
'@param layoutName Optional String. Empty reads the current state; a name reads
'  that saved layout.
'@param writeSizes Optional Boolean. True writes each stored size to the sheet.
'  Default True.
'@return Long. The number of rows that matched an entry.
'@export
Public Function Load(ByVal entries As ShowHide, _
                     Optional ByVal layout As ShowHideLayout, _
                     Optional ByVal layoutName As String = vbNullString, _
                     Optional ByVal writeSizes As Boolean = True) As Long
    Dim tableRows As Variant
    Dim token As String
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim hiddenCol As Long
    Dim sizeCol As Long
    Dim orientCol As Long
    Dim layoutCol As Long
    Dim counter As Long
    Dim entryIndex As Long
    Dim position As Long
    Dim applied As Long
    Dim entrySize As Double
    Dim orientation As String

    If entries Is Nothing Then Exit Function
    If this.lo Is Nothing Then Exit Function
    If this.lo.DataBodyRange Is Nothing Then Exit Function

    layerCol = ColumnIndex(COL_LAYER)
    fieldCol = ColumnIndex(COL_FIELD_KEY)
    hiddenCol = ColumnIndex(COL_HIDDEN)
    sizeCol = SizeColumnIndex()
    orientCol = ColumnIndex(COL_ORIENTATION)
    layoutCol = ColumnIndex(COL_LAYOUT)

    If layerCol = 0 Or fieldCol = 0 Or hiddenCol = 0 Then Exit Function

    layoutName = Trim$(layoutName)
    token = LayerToken(entries.TargetLayer)
    tableRows = this.lo.DataBodyRange.Value

    If Not layout Is Nothing Then layout.BeginBatch

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If StrComp(CStr(tableRows(counter, layerCol)), token, vbTextCompare) = 0 _
           And LayoutMatches(RowLayoutName(tableRows, counter, layoutCol), layoutName) Then
            entryIndex = entries.IndexOf(CStr(tableRows(counter, fieldCol)))

            If entryIndex > 0 Then
                applied = applied + 1
                entries.SetHidden entryIndex, ToBool(tableRows(counter, hiddenCol))

                If Not layout Is Nothing Then
                    position = entries.PositionIndex(entryIndex)

                    If position > 0 And sizeCol > 0 And writeSizes Then
                        entrySize = ToSize(tableRows(counter, sizeCol))
                        If entrySize > 0 Then layout.SetSize position, entrySize
                    End If

                    If position > 0 And orientCol > 0 And layout.SupportsOrientation Then
                        orientation = LCase$(Trim$(CStr(tableRows(counter, orientCol))))
                        If LenB(orientation) > 0 Then
                            layout.SetOrientation position, (orientation = "vertical")
                        End If
                    End If
                End If
            End If
        End If
    Next counter

    If Not layout Is Nothing Then layout.EndBatch

    'The list now says what the current state holds; a named layout is not it
    If applied > 0 And LenB(layoutName) = 0 Then entries.MarkStored

    Load = applied
End Function


'@section Saved layouts
'===============================================================================
'@description The named copies of the show/hide state, and their bookkeeping.

'@label:max-saved-layouts
'@prop-title How many saved layouts one store holds at most
'@export
Public Property Get MaxSavedLayouts() As Long
    MaxSavedLayouts = MAX_SAVED_LAYOUTS
End Property

'@label:layout-names
'@sub-title The saved layout names, in the order they first appear
'@details
'Names are compared without regard to case, and each keeps the casing of its
'first row. The current state has no name and is never listed.
'@return BetterArray. The distinct names, empty when nothing is saved.
'@export
Public Function LayoutNames() As BetterArray
    Dim names As BetterArray
    Dim seen As Collection
    Dim tableRows As Variant
    Dim layoutCol As Long
    Dim counter As Long
    Dim layoutName As String

    Set names = New BetterArray
    Set LayoutNames = names

    If this.lo Is Nothing Then Exit Function
    If this.lo.DataBodyRange Is Nothing Then Exit Function

    layoutCol = ColumnIndex(COL_LAYOUT)
    If layoutCol = 0 Then Exit Function

    Set seen = New Collection
    tableRows = this.lo.DataBodyRange.Value

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        layoutName = Trim$(CStr(tableRows(counter, layoutCol)))
        If LenB(layoutName) > 0 Then
            'A Collection has no "does this key exist" test, so the miss is
            'read off Err, the same shape EventLinelist uses for its stores.
            On Error Resume Next
                Err.Clear
                seen.Add layoutName, LCase$(layoutName)
                If Err.Number = 0 Then names.Push layoutName
                Err.Clear
            On Error GoTo 0
        End If
    Next counter
End Function

'@label:layout-count
'@sub-title How many saved layouts the store holds
'@export
Public Function LayoutCount() As Long
    LayoutCount = LayoutNames().Length
End Function

'@label:has-layout
'@sub-title Whether a saved layout of one name is stored
'@param layoutName String. The name to look for, compared without regard to case.
'@return Boolean. True when at least one row carries the name.
'@export
Public Function HasLayout(ByVal layoutName As String) As Boolean
    Dim tableRows As Variant
    Dim layoutCol As Long
    Dim counter As Long

    layoutName = Trim$(layoutName)
    If LenB(layoutName) = 0 Then Exit Function
    If this.lo Is Nothing Then Exit Function
    If this.lo.DataBodyRange Is Nothing Then Exit Function

    layoutCol = ColumnIndex(COL_LAYOUT)
    If layoutCol = 0 Then Exit Function

    tableRows = this.lo.DataBodyRange.Value

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If LayoutMatches(RowLayoutName(tableRows, counter, layoutCol), layoutName) Then
            HasLayout = True
            Exit Function
        End If
    Next counter
End Function

'@label:delete-layout
'@sub-title Drop every row of one saved layout, over all its layers
'@param layoutName String. The name to drop. The current state has no name and
'  is cleared through Clear.
'@throws ProjectError.ElementNotFound When the store has no table.
'@throws ProjectError.InvalidArgument When the name is empty.
'@export
Public Sub DeleteLayout(ByVal layoutName As String)
    Dim tableRows As Variant
    Dim kept As Variant
    Dim keptCount As Long
    Dim layoutCol As Long
    Dim counter As Long
    Dim target As Long
    Dim colCount As Long

    RequireTable

    layoutName = Trim$(layoutName)
    If LenB(layoutName) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "A name is needed to delete a saved layout"
    End If

    layoutCol = ColumnIndex(COL_LAYOUT)
    If layoutCol = 0 Then Exit Sub
    If this.lo.DataBodyRange Is Nothing Then Exit Sub

    colCount = this.lo.ListColumns.Count
    tableRows = this.lo.DataBodyRange.Value
    ReDim kept(1 To UBound(tableRows, 1), 1 To colCount)

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If Not LayoutMatches(RowLayoutName(tableRows, counter, layoutCol), layoutName) Then
            keptCount = keptCount + 1
            For target = 1 To colCount
                kept(keptCount, target) = tableRows(counter, target)
            Next target
        End If
    Next counter

    WriteAll kept, keptCount
End Sub

'@label:rename-layout
'@sub-title Give one saved layout another name
'@details
'Renaming onto a name the store already holds is refused, because the rows of
'the two layouts would merge and neither could be told apart again. Renaming a
'layout onto its own name with other casing is allowed.
'@param oldName String. The name the layout holds.
'@param newName String. The name to give it.
'@throws ProjectError.ElementNotFound When the store has no table or no layout
'  of the old name.
'@throws ProjectError.InvalidArgument When a name is empty or the new name is
'  already stored.
'@export
Public Sub RenameLayout(ByVal oldName As String, ByVal newName As String)
    Dim tableRows As Variant
    Dim layoutCol As Long
    Dim counter As Long
    Dim sameName As Boolean

    RequireTable

    oldName = Trim$(oldName)
    newName = Trim$(newName)

    If LenB(oldName) = 0 Or LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "Renaming a saved layout needs the old name and the new one"
    End If

    If Not HasLayout(oldName) Then
        ThrowError ProjectError.ElementNotFound, _
                   "There is no saved layout named '" & oldName & "'"
    End If

    sameName = LayoutMatches(oldName, newName)
    If Not sameName Then
        If HasLayout(newName) Then
            ThrowError ProjectError.InvalidArgument, _
                       "There is already a saved layout named '" & newName & "'"
        End If
    End If

    layoutCol = ColumnIndex(COL_LAYOUT)
    tableRows = this.lo.DataBodyRange.Value

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If LayoutMatches(RowLayoutName(tableRows, counter, layoutCol), oldName) Then
            tableRows(counter, layoutCol) = newName
        End If
    Next counter

    this.lo.DataBodyRange.Value = tableRows
End Sub

'@label:merge-layouts-from
'@sub-title Copy another store's saved layouts into this one
'@details
'How saved layouts travel: the exporter merges the linelist's layouts into the
'migration file's store, and the importer merges the file's layouts back. A
'name both stores hold is taken from the source, replacing the local rows. A
'new name is taken while the cap has room; the rest are skipped and answered
'to the caller, so an import can report them. Rows are copied as they stand -
'a variable the receiving dictionary does not carry is skipped by Load later.
'@param source ShowHideStore. The store to copy from.
'@return BetterArray. The names that were skipped over the cap, empty when
'  every layout came across.
'@export
Public Function MergeLayoutsFrom(ByVal source As ShowHideStore) As BetterArray
    Dim skipped As BetterArray
    Dim accepted As Collection
    Dim incoming As BetterArray
    Dim srcRows As Variant
    Dim tableRows As Variant
    Dim block As Variant
    Dim srcLo As ListObject
    Dim total As Long
    Dim counter As Long
    Dim target As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim incomingName As String
    Dim rowName As String
    Dim layoutCol As Long
    Dim srcLayoutCol As Long

    Set skipped = New BetterArray
    Set MergeLayoutsFrom = skipped

    RequireTable
    layoutCol = ColumnIndex(COL_LAYOUT)
    If layoutCol = 0 Then Exit Function

    If source Is Nothing Then Exit Function
    If Not source.HasTable Then Exit Function

    Set srcLo = source.Table
    srcLayoutCol = ColumnIndexOf(srcLo, COL_LAYOUT)
    If srcLayoutCol = 0 Then Exit Function
    If srcLo.DataBodyRange Is Nothing Then Exit Function

    Set incoming = source.LayoutNames()
    If incoming.Length = 0 Then Exit Function

    'Which incoming names come across. A collision costs no slot, because the
    'local rows of that name go first.
    Set accepted = New Collection
    total = LayoutCount()

    For counter = incoming.LowerBound To incoming.UpperBound
        incomingName = CStr(incoming.Item(counter))
        If HasLayout(incomingName) Then
            accepted.Add incomingName, LCase$(incomingName)
        ElseIf total < MAX_SAVED_LAYOUTS Then
            accepted.Add incomingName, LCase$(incomingName)
            total = total + 1
        Else
            skipped.Push incomingName
        End If
    Next counter

    If accepted.Count = 0 Then Exit Function

    srcRows = srcLo.DataBodyRange.Value
    colCount = this.lo.ListColumns.Count

    If this.lo.DataBodyRange Is Nothing Then
        ReDim block(1 To UBound(srcRows, 1), 1 To colCount)
    Else
        tableRows = this.lo.DataBodyRange.Value
        ReDim block(1 To UBound(tableRows, 1) + UBound(srcRows, 1), 1 To colCount)

        'Local rows first, minus the layouts the source replaces
        For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
            rowName = RowLayoutName(tableRows, counter, layoutCol)
            If LenB(rowName) = 0 Or Not InCollection(accepted, rowName) Then
                rowCount = rowCount + 1
                For target = 1 To colCount
                    block(rowCount, target) = tableRows(counter, target)
                Next target
            End If
        Next counter
    End If

    'Then the accepted rows of the source, mapped column by column
    rowCount = AppendSourceLayoutRows(block, rowCount, srcLo, srcRows, accepted, srcLayoutCol)

    WriteAll block, rowCount
End Function


'@section Helpers
'===============================================================================
'@description Private workers of the three factories and the two writers.

'@label:provision
'@sub-title Find the store table on a sheet, or build it
'@param sh Worksheet. The sheet to provision.
'@return ListObject. The store table.
Private Function Provision(ByVal sh As Worksheet) As ListObject
    Dim lo As ListObject
    Dim headerRange As Range
    Dim headers As Variant
    Dim counter As Long

    DropLegacyTables sh

    Set lo = NamedTable(sh)
    If Not lo Is Nothing Then
        EnsureLayoutColumn lo
        Set Provision = lo
        Exit Function
    End If

    headers = Array(COL_LAYER, COL_FIELD_KEY, COL_HEADER, COL_HIDDEN, _
                    COL_SIZE, COL_ORIENTATION, COL_LAYOUT)

    For counter = LBound(headers) To UBound(headers)
        sh.Cells(1, counter - LBound(headers) + 1).Value = CStr(headers(counter))
    Next counter

    Set headerRange = sh.Range(sh.Cells(1, 1), sh.Cells(1, COLUMN_COUNT))
    Set lo = sh.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
    lo.Name = STORE_TABLE

    Set Provision = lo
End Function

'@label:ensure-layout-column
'@sub-title Add the layout_name column to a table an older build made
'@details
'A linelist generated before saved layouts existed carries a six column table.
'The column lands at the right edge and every existing row keeps an empty
'cell there, which is exactly what a current-state row looks like.
'@param lo ListObject. The table to check.
Private Sub EnsureLayoutColumn(ByVal lo As ListObject)
    Dim found As Range

    On Error Resume Next
    Set found = lo.HeaderRowRange.Find(What:=COL_LAYOUT, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If Not found Is Nothing Then Exit Sub

    On Error Resume Next
    lo.ListColumns.Add.Name = COL_LAYOUT
    On Error GoTo 0
End Sub

'@label:named-table
'@sub-title Look for a store table by name
'@details
'The name this class writes comes first, then the name an older build wrote.
'A table under any other name is left alone, so provisioning never adopts a
'table that holds something else.
'@param sh Worksheet. The sheet to search.
'@return ListObject. The table, or Nothing.
Private Function NamedTable(ByVal sh As Worksheet) As ListObject
    Dim lo As ListObject

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Set lo = sh.ListObjects(STORE_TABLE)
    If lo Is Nothing Then Set lo = sh.ListObjects(LEGACY_TABLE)
    On Error GoTo 0

    Set NamedTable = lo
End Function

'@label:find-table
'@sub-title Look for a store table already on a sheet
'@details
'Either of the two known names first, then whatever single table the sheet
'carries. The last step is what lets a migration file written by a build this
'one has never heard of still be read.
'@param sh Worksheet. The sheet to search.
'@return ListObject. The table, or Nothing.
Private Function FindTable(ByVal sh As Worksheet) As ListObject
    Dim lo As ListObject

    If sh Is Nothing Then Exit Function

    Set lo = NamedTable(sh)
    If lo Is Nothing And sh.ListObjects.Count > 0 Then
        Set lo = sh.ListObjects(1)
    End If

    Set FindTable = lo
End Function

'@label:drop-legacy-tables
'@sub-title Delete the per-sheet tables of the first show/hide design
'@param sh Worksheet. The sheet to sweep.
Private Sub DropLegacyTables(ByVal sh As Worksheet)
    Dim counter As Long
    Dim lo As ListObject

    For counter = sh.ListObjects.Count To 1 Step -1
        Set lo = sh.ListObjects(counter)
        If InStr(1, lo.Name, LEGACY_PREFIX, vbTextCompare) = 1 Then
            On Error Resume Next
            lo.Range.Delete Shift:=xlToLeft
            On Error GoTo 0
        End If
    Next counter
End Sub

'@label:kept-rows
'@sub-title Every row a save or a clear carries over untouched
'@details
'A row is replaced when its layer, its layout name and its field key all
'match. With no entry list, every key of the layer and state matches, which is
'what Clear asks for.
'@param token String. The layer token being replaced.
'@param layerCol Long. The index of the layer column.
'@param layoutName String. The state being replaced, empty for the current one.
'@param layoutCol Long. The index of the layout column, or 0 on an old table.
'@param entries ShowHide. The entry list whose keys are replaced, or Nothing.
'@param fieldCol Long. The index of the field key column, or 0.
'@param keptCount Long. Set to the number of rows kept.
'@return Variant. A two dimensional block of the kept rows, or Empty.
Private Function KeptRows(ByVal token As String, _
                          ByVal layerCol As Long, _
                          ByVal layoutName As String, _
                          ByVal layoutCol As Long, _
                          ByVal entries As ShowHide, _
                          ByVal fieldCol As Long, _
                          ByRef keptCount As Long) As Variant
    Dim tableRows As Variant
    Dim kept As Variant
    Dim counter As Long
    Dim colCount As Long
    Dim target As Long
    Dim replaced As Boolean
    Dim sizeCol As Long
    Dim oldSize As Double

    keptCount = 0

    If this.lo.DataBodyRange Is Nothing Then Exit Function
    If layerCol = 0 Then Exit Function

    'While a save holds the collection, the size of every replaced row is filed
    'in it, keyed on the lower case field key, so the save can keep a size the
    'layout cannot measure.
    If Not this.droppedSizes Is Nothing Then sizeCol = SizeColumnIndex()

    colCount = this.lo.ListColumns.Count
    tableRows = this.lo.DataBodyRange.Value
    ReDim kept(1 To UBound(tableRows, 1), 1 To colCount)

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        replaced = (StrComp(CStr(tableRows(counter, layerCol)), token, vbTextCompare) = 0)

        If replaced Then
            replaced = LayoutMatches(RowLayoutName(tableRows, counter, layoutCol), layoutName)
        End If

        If replaced And Not (entries Is Nothing) And fieldCol > 0 Then
            replaced = entries.HasField(CStr(tableRows(counter, fieldCol)))
        End If

        If Not replaced Then
            keptCount = keptCount + 1
            For target = 1 To colCount
                kept(keptCount, target) = tableRows(counter, target)
            Next target
        ElseIf sizeCol > 0 And fieldCol > 0 Then
            oldSize = ToSize(tableRows(counter, sizeCol))
            If oldSize > 0 Then
                On Error Resume Next
                this.droppedSizes.Add oldSize, LCase$(Trim$(CStr(tableRows(counter, fieldCol))))
                On Error GoTo 0
            End If
        End If
    Next counter

    KeptRows = kept
End Function

'@label:dropped-size
'@sub-title The size a replaced row held for one field, or 0
'@param fieldKey String. The field to look up.
'@return Double. The stored size, or 0 when the row held none.
Private Function DroppedSize(ByVal fieldKey As String) As Double
    If this.droppedSizes Is Nothing Then Exit Function

    On Error Resume Next
    DroppedSize = this.droppedSizes.Item(LCase$(Trim$(fieldKey)))
    On Error GoTo 0
End Function

'@label:append-source-layout-rows
'@sub-title Copy the accepted layout rows of another table into a block
'@details
'Columns are matched by header on both sides, so a source written by another
'build still lands in the right cells here.
'@param block Variant. The block being filled, sized by the caller.
'@param startCount Long. How many rows the block already holds.
'@param srcLo ListObject. The source table, for its headers.
'@param srcRows Variant. The source body, already read.
'@param accepted Collection. The layout names to copy, keyed on their lower case.
'@param srcLayoutCol Long. The source's layout column index.
'@return Long. The row count of the block once the copies are in.
Private Function AppendSourceLayoutRows(ByRef block As Variant, _
                                        ByVal startCount As Long, _
                                        ByVal srcLo As ListObject, _
                                        ByVal srcRows As Variant, _
                                        ByVal accepted As Collection, _
                                        ByVal srcLayoutCol As Long) As Long
    Dim rowCount As Long
    Dim counter As Long
    Dim rowName As String
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim headerCol As Long
    Dim hiddenCol As Long
    Dim sizeCol As Long
    Dim orientCol As Long
    Dim layoutCol As Long
    Dim srcLayerCol As Long
    Dim srcFieldCol As Long
    Dim srcHeaderCol As Long
    Dim srcHiddenCol As Long
    Dim srcSizeCol As Long
    Dim srcOrientCol As Long

    layerCol = ColumnIndex(COL_LAYER)
    fieldCol = ColumnIndex(COL_FIELD_KEY)
    headerCol = ColumnIndex(COL_HEADER)
    hiddenCol = ColumnIndex(COL_HIDDEN)
    sizeCol = SizeColumnIndex()
    orientCol = ColumnIndex(COL_ORIENTATION)
    layoutCol = ColumnIndex(COL_LAYOUT)

    srcLayerCol = ColumnIndexOf(srcLo, COL_LAYER)
    srcFieldCol = ColumnIndexOf(srcLo, COL_FIELD_KEY)
    srcHeaderCol = ColumnIndexOf(srcLo, COL_HEADER)
    srcHiddenCol = ColumnIndexOf(srcLo, COL_HIDDEN)
    srcSizeCol = ColumnIndexOf(srcLo, COL_SIZE)
    If srcSizeCol = 0 Then srcSizeCol = ColumnIndexOf(srcLo, COL_LEGACY_SIZE)
    srcOrientCol = ColumnIndexOf(srcLo, COL_ORIENTATION)

    rowCount = startCount
    If srcLayerCol = 0 Or srcFieldCol = 0 Or srcHiddenCol = 0 Then
        AppendSourceLayoutRows = rowCount
        Exit Function
    End If

    For counter = LBound(srcRows, 1) To UBound(srcRows, 1)
        rowName = RowLayoutName(srcRows, counter, srcLayoutCol)

        If LenB(rowName) > 0 Then
            If InCollection(accepted, rowName) Then
                rowCount = rowCount + 1
                block(rowCount, layerCol) = srcRows(counter, srcLayerCol)
                block(rowCount, fieldCol) = srcRows(counter, srcFieldCol)
                If headerCol > 0 And srcHeaderCol > 0 Then _
                    block(rowCount, headerCol) = srcRows(counter, srcHeaderCol)
                block(rowCount, hiddenCol) = srcRows(counter, srcHiddenCol)
                If sizeCol > 0 And srcSizeCol > 0 Then _
                    block(rowCount, sizeCol) = srcRows(counter, srcSizeCol)
                If orientCol > 0 And srcOrientCol > 0 Then _
                    block(rowCount, orientCol) = srcRows(counter, srcOrientCol)
                block(rowCount, layoutCol) = rowName
            End If
        End If
    Next counter

    AppendSourceLayoutRows = rowCount
End Function

'@label:row-layout-name
'@sub-title The layout name one row carries
'@param tableRows Variant. The body block being read.
'@param rowIndex Long. The row to read.
'@param layoutCol Long. The layout column index, or 0 on an old table.
'@return String. The trimmed name, empty for a current-state row.
Private Function RowLayoutName(ByVal tableRows As Variant, _
                               ByVal rowIndex As Long, _
                               ByVal layoutCol As Long) As String
    If layoutCol = 0 Then Exit Function
    RowLayoutName = Trim$(CStr(tableRows(rowIndex, layoutCol)))
End Function

'@label:layout-matches
'@sub-title Whether two layout names are the same name
'@param firstName String. One name, already trimmed.
'@param secondName String. The other, already trimmed.
'@return Boolean. True when they match without regard to case.
Private Function LayoutMatches(ByVal firstName As String, _
                               ByVal secondName As String) As Boolean
    LayoutMatches = (StrComp(firstName, secondName, vbTextCompare) = 0)
End Function

'@label:in-collection
'@sub-title Whether a keyed Collection holds one name
'@param keyed Collection. The collection, keyed on lower cased names.
'@param layoutName String. The name to look for.
'@return Boolean. True when the key is there.
Private Function InCollection(ByVal keyed As Collection, _
                              ByVal layoutName As String) As Boolean
    On Error Resume Next
        Err.Clear
        keyed.Item LCase$(layoutName)
        InCollection = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0
End Function

'@label:write-all
'@sub-title Replace the whole table body with one block
'@details
'One write and one resize, whatever the row count. Writing row by row through
'ListRows.Add costs one Excel crossing per row, and a data sheet with two
'hundred variables closes its form four times slower for it.
'@param block Variant. The rows to write, one dimension per row and column.
'@param rowCount Long. How many rows of the block to keep.
Private Sub WriteAll(ByVal block As Variant, ByVal rowCount As Long)
    Dim colCount As Long
    Dim firstCell As Range
    Dim target As Range

    colCount = this.lo.ListColumns.Count
    Set firstCell = this.lo.HeaderRowRange.Cells(1, 1).Offset(1)

    'Clear whatever the table held before, body and any leftovers below it
    If Not this.lo.DataBodyRange Is Nothing Then this.lo.DataBodyRange.ClearContents

    If rowCount = 0 Then
        On Error Resume Next
        this.lo.Resize this.lo.HeaderRowRange
        On Error GoTo 0
        Exit Sub
    End If

    Set target = firstCell.Resize(rowCount, colCount)
    this.lo.Resize this.lo.HeaderRowRange.Resize(rowCount + 1, colCount)
    target.Value = TrimBlock(block, rowCount, colCount)
End Sub

'@label:trim-block
'@sub-title Cut a block down to the rows that are used
'@param block Variant. The oversized block.
'@param rowCount Long. How many rows to keep.
'@param colCount Long. How many columns to keep.
'@return Variant. A block of exactly rowCount by colCount.
Private Function TrimBlock(ByVal block As Variant, _
                           ByVal rowCount As Long, _
                           ByVal colCount As Long) As Variant
    Dim trimmed As Variant
    Dim rowIdx As Long
    Dim colIdx As Long

    If UBound(block, 1) = rowCount And UBound(block, 2) = colCount Then
        TrimBlock = block
        Exit Function
    End If

    ReDim trimmed(1 To rowCount, 1 To colCount)
    For rowIdx = 1 To rowCount
        For colIdx = 1 To colCount
            trimmed(rowIdx, colIdx) = block(rowIdx, colIdx)
        Next colIdx
    Next rowIdx

    TrimBlock = trimmed
End Function

'@label:column-index
'@sub-title Where one named column sits in the store's own table
'@param headerName String. The column header to find.
'@return Long. The one based column number, or 0 when the header is absent.
Private Function ColumnIndex(ByVal headerName As String) As Long
    ColumnIndex = ColumnIndexOf(this.lo, headerName)
End Function

'@label:column-index-of
'@sub-title Where one named column sits in any store table
'@details
'The merge reads another store's table through this, so both sides are matched
'by header and a source written by another build still maps.
'@param lo ListObject. The table to search.
'@param headerName String. The column header to find.
'@return Long. The one based column number, or 0 when the header is absent.
Private Function ColumnIndexOf(ByVal lo As ListObject, _
                               ByVal headerName As String) As Long
    Dim found As Range

    If lo Is Nothing Then Exit Function

    On Error Resume Next
    Set found = lo.HeaderRowRange.Find(What:=headerName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If found Is Nothing Then Exit Function
    ColumnIndexOf = found.Column - lo.HeaderRowRange.Column + 1
End Function

'@label:size-column-index
'@sub-title Where the size column sits, under either of its two names
'@details
'A file written by an older build calls it `column_width`.
'@return Long. The one based column number, or 0 when neither name is there.
Private Function SizeColumnIndex() As Long
    Dim answer As Long

    answer = ColumnIndex(COL_SIZE)
    If answer = 0 Then answer = ColumnIndex(COL_LEGACY_SIZE)
    SizeColumnIndex = answer
End Function

'@label:require-table
'@sub-title Refuse to write when there is no table
'@throws ProjectError.ElementNotFound When the store is bound to nothing.
Private Sub RequireTable()
    If this.lo Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The show/hide store has no table to write to"
    End If
End Sub

'@label:layer-token
'@sub-title The word one layer is written under
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to name.
'@return String. The token stored in the layer column.
'@throws ProjectError.InvalidArgument When the layer is unknown.
Private Function LayerToken(ByVal layer As Byte) As String
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
                       "There is no show/hide layer '" & CStr(layer) & "' in a linelist"
    End Select
End Function

'@label:orientation-token
'@sub-title The word a header direction is written under
'@param vertical Boolean. True when the header is turned.
'@return String. "vertical" or "horizontal".
Private Function OrientationToken(ByVal vertical As Boolean) As String
    If vertical Then
        OrientationToken = "vertical"
    Else
        OrientationToken = "horizontal"
    End If
End Function

'@label:bool-to-string
'@sub-title Write a hidden flag in a form no language changes
'@param value Boolean. The flag to write.
'@return String. "true" or "false".
Private Function BoolToString(ByVal value As Boolean) As String
    If value Then
        BoolToString = "true"
    Else
        BoolToString = "false"
    End If
End Function

'@label:to-bool
'@sub-title Read a stored hidden flag
'@param value Variant. The cell value.
'@return Boolean. True when the cell says so.
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

'@label:to-size
'@sub-title Read a stored size, answering zero for anything unusable
'@param value Variant. The cell value.
'@return Double. The size, or 0.
Private Function ToSize(ByVal value As Variant) As Double
    Dim answer As Double

    If IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    On Error Resume Next
    answer = CDbl(value)
    On Error GoTo 0

    ToSize = answer
End Function

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
