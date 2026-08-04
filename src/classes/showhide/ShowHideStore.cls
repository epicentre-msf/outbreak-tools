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
'four layers, told apart by the `layer` column. Save writes the choices of one
'layer, Load reads them back, Clear empties the table, HasLayer says whether a
'layer has rows, and HasTable, Table and RowCount answer where the store
'stands.
'
'THE STORE MAKES ITS OWN TABLE
'-------------------------------------------------------------------------------
'Create and CreateOnSheet build the ListObject when it is absent.
'
'CreateForRead binds to whatever table a workbook already carries and builds
'nothing, which is what reading a migration file needs.
'
'SIX COLUMNS
'-------------------------------------------------------------------------------
'  layer         hlist / printed / vlist / crf
'  field_key     the variable name, and the key rows are matched on
'  header_text   for a person reading the sheet
'  hidden_flag   true / false
'  entry_size    the width or the height the entry has when shown, blank when
'                unknown
'  orientation   vertical / horizontal, blank off the printed layer
'
'The size and the hidden flag are two separate columns. A size at or below zero
'means "leave the sheet alone".
'
'A file written by an older build carries `column_width` where this one writes
'`entry_size`. Load reads either name.
'@depends ShowHide, ShowHideLayout, Checking
'@version 1.0 (2026-08-01)


Private Const CLASS_NAME As String = "ShowHideStore"

'The worksheet a linelist keeps its show/hide table on
Private Const STORE_SHEET As String = "show_hide__"

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
Private Const COL_LEGACY_SIZE As String = "column_width"

Private Const COLUMN_COUNT As Long = 6

Private Type TShowHideStore
    lo As ListObject
    sh As Worksheet
End Type

Private this As TShowHideStore


'@section Instantiation
'===============================================================================
'@description Three ways in: two that provision, one that reads what is there.

'@label:create
'@sub-title Open the store of a workbook, making the table when it is absent
'@param wkb Workbook. The workbook holding the show_hide__ worksheet.
'@return ShowHideStore. A store ready to save and load.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@throws ProjectError.ElementNotFound When the workbook has no show_hide__
'  worksheet.
'@export
Public Function Create(ByVal wkb As Workbook) As ShowHideStore
    Dim sh As Worksheet

    If wkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A workbook is needed to find the show/hide store"
    End If

    On Error Resume Next
    Set sh = wkb.Worksheets(STORE_SHEET)
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
'@sub-title Write one layer's choices, leaving the other layers alone
'@details
'The rows of the entry list's own layer are replaced and every other layer's rows are
'carried over untouched, so four layers share one table.
'
'Without a layout the store writes visibility alone, which is what closing the
'form needs. With one it also records the size each entry has when shown and,
'on a printed sheet, the direction of its header - which is what a migration
'export needs.
'@param entries ShowHide. The entry list to write.
'@param layout Optional ShowHideLayout. The sheet to read sizes from.
'@throws ProjectError.ObjectNotInitialized When the entry list is Nothing.
'@throws ProjectError.ElementNotFound When the store has no table.
'@export
Public Sub Save(ByVal entries As ShowHide, Optional ByVal layout As ShowHideLayout)
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

    If layerCol = 0 Or fieldCol = 0 Or hiddenCol = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The show/hide table is missing one of layer, field_key or hidden_flag"
    End If

    token = LayerToken(entries.TargetLayer)
    kept = RowsOfOtherLayers(token, layerCol, keptCount)

    If keptCount + entries.EntryCount = 0 Then
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

        If Not layout Is Nothing Then
            position = entries.PositionIndex(idx)
            If position > 0 Then
                If sizeCol > 0 Then
                    entrySize = layout.SizeWhenShown(position)
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
End Sub

'@label:clear
'@sub-title Drop every row of one layer
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to drop.
'@throws ProjectError.ElementNotFound When the store has no table.
'@export
Public Sub Clear(ByVal layer As Byte)
    Dim kept As Variant
    Dim keptCount As Long
    Dim layerCol As Long

    RequireTable

    layerCol = ColumnIndex(COL_LAYER)
    If layerCol = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The show/hide table is missing its " & COL_LAYER & " column"
    End If

    kept = RowsOfOtherLayers(LayerToken(layer), layerCol, keptCount)

    WriteAll kept, keptCount
End Sub


'@section Loading
'===============================================================================
'@description Reading the table back into an entry list.

'@label:load
'@sub-title Read one layer's saved choices into an entry list
'@details
'Only the rows of the entry list's own layer are read. A variable the list does not
'carry is skipped, which is what happens when the dictionary changed between two
'sessions.
'
'With a layout, the saved size and header direction are applied to the sheet as
'well. Visibility is left to ShowHide.Apply, so a caller that wants the sheet to
'match the list calls Load and then Apply.
'@param entries ShowHide. The entry list to fill.
'@param layout Optional ShowHideLayout. The sheet to size.
'@return Long. The number of rows that matched an entry.
'@export
Public Function Load(ByVal entries As ShowHide, Optional ByVal layout As ShowHideLayout) As Long
    Dim tableRows As Variant
    Dim token As String
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim hiddenCol As Long
    Dim sizeCol As Long
    Dim orientCol As Long
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

    If layerCol = 0 Or fieldCol = 0 Or hiddenCol = 0 Then Exit Function

    token = LayerToken(entries.TargetLayer)
    tableRows = this.lo.DataBodyRange.Value

    If Not layout Is Nothing Then layout.BeginBatch

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If StrComp(CStr(tableRows(counter, layerCol)), token, vbTextCompare) = 0 Then
            entryIndex = entries.IndexOf(CStr(tableRows(counter, fieldCol)))

            If entryIndex > 0 Then
                applied = applied + 1
                entries.SetHidden entryIndex, ToBool(tableRows(counter, hiddenCol))

                If Not layout Is Nothing Then
                    position = entries.PositionIndex(entryIndex)

                    If position > 0 And sizeCol > 0 Then
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

    Load = applied
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
        Set Provision = lo
        Exit Function
    End If

    headers = Array(COL_LAYER, COL_FIELD_KEY, COL_HEADER, COL_HIDDEN, _
                    COL_SIZE, COL_ORIENTATION)

    For counter = LBound(headers) To UBound(headers)
        sh.Cells(1, counter - LBound(headers) + 1).Value = CStr(headers(counter))
    Next counter

    Set headerRange = sh.Range(sh.Cells(1, 1), sh.Cells(1, COLUMN_COUNT))
    Set lo = sh.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
    lo.Name = STORE_TABLE

    Set Provision = lo
End Function

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

'@label:rows-of-other-layers
'@sub-title Every row that does not belong to one layer
'@param token String. The layer token being replaced.
'@param layerCol Long. The index of the layer column.
'@param keptCount Long. Set to the number of rows kept.
'@return Variant. A two dimensional block of the kept rows, or Empty.
Private Function RowsOfOtherLayers(ByVal token As String, _
                                   ByVal layerCol As Long, _
                                   ByRef keptCount As Long) As Variant
    Dim tableRows As Variant
    Dim kept As Variant
    Dim counter As Long
    Dim colCount As Long
    Dim target As Long

    keptCount = 0

    If this.lo.DataBodyRange Is Nothing Then Exit Function
    If layerCol = 0 Then Exit Function

    colCount = this.lo.ListColumns.Count
    tableRows = this.lo.DataBodyRange.Value
    ReDim kept(1 To UBound(tableRows, 1), 1 To colCount)

    For counter = LBound(tableRows, 1) To UBound(tableRows, 1)
        If StrComp(CStr(tableRows(counter, layerCol)), token, vbTextCompare) <> 0 Then
            keptCount = keptCount + 1
            For target = 1 To colCount
                kept(keptCount, target) = tableRows(counter, target)
            Next target
        End If
    Next counter

    RowsOfOtherLayers = kept
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
'@sub-title Where one named column sits in the table
'@param headerName String. The column header to find.
'@return Long. The one based column number, or 0 when the header is absent.
Private Function ColumnIndex(ByVal headerName As String) As Long
    Dim found As Range

    If this.lo Is Nothing Then Exit Function

    On Error Resume Next
    Set found = this.lo.HeaderRowRange.Find(What:=headerName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If found Is Nothing Then Exit Function
    ColumnIndex = found.Column - this.lo.HeaderRowRange.Column + 1
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
