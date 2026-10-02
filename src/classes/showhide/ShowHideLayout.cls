VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Where show/hide entries sit on one worksheet and how they look"

'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Where show/hide entries sit on one worksheet and how they look")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'The two directions a show/hide entry can occupy on a sheet.
Public Enum ShowHideAxis
    ShowHideAxisColumn = 1
    ShowHideAxisRow = 2
End Enum


'@class ShowHideLayout
'@description
'Everything show/hide does to a worksheet. One instance covers one sheet of one
'layer. The entry list, ShowHide, says which variables are hidden; this class
'knows where each of them sits and writes the change. IsHidden and SetHidden
'read and write the hidden state, Size, SetSize and SizeWhenShown the width or
'height, IsVertical and SetOrientation the orientation of a printed entry, and
'FailureCount counts the writes that Excel refused.
'
'THE AXIS IS DECIDED HERE AND NOWHERE ELSE
'-------------------------------------------------------------------------------
'  HList, Printed   a column
'  VList, CRF       a row
'
'A CRF holds one variable per row and reads its position from `crf index`, its
'own dictionary column.
'
'PROTECTION IS BRACKETED INSIDE THE CLASS
'-------------------------------------------------------------------------------
'Every method that writes to the sheet unprotects it first and protects it
'again after. Reads leave the protection alone. A loop that writes many
'positions calls BeginBatch first and EndBatch at the end, which holds one
'bracket open over the whole run.
'
'A SIZE OF ZERO IS NEVER WRITTEN
'-------------------------------------------------------------------------------
'Excel reports a hidden column as width 0 and hides a column that is set to
'width 0. SetSize refuses anything at or below zero. Hiding is SetHidden's
'work, and a blank size cell leaves the sheet alone.
'
'A HIDDEN POSITION REMEMBERS ITS SIZE
'-------------------------------------------------------------------------------
'SetHidden reads the size before it hides and keeps it on the instance, keyed
'on the position. SizeWhenShown answers from that and shows a column only when
'it holds nothing for it. WriteCount says how many writes went out since the
'bind or the last ResetWriteCount, so a save can tell a session that moved
'nothing from one that did.
'@depends Passwords, Checking
'@version 1.1 (2026-09-01)


Private Const CLASS_NAME As String = "ShowHideLayout"

'The named cell a printed sheet puts at the head of its table
Private Const PRINT_ANCHOR_SUFFIX As String = "_PRINTSTART"

'How far above the anchor the printed header row sits. PRINTSTART is on the
'first line of the table, so the header the user reads is the line above it.
Private Const PRINT_HEADER_OFFSET As Long = -1

'The row height a turned header needs to stay readable
Private Const VERTICAL_HEADER_HEIGHT As Double = 100

Private Type TShowHideLayout
    sh As Worksheet
    layer As Byte
    axisValue As Byte
    pass As Passwords
    baseTableName As String
    lockDepth As Long
    failures As Long
    writes As Long                  'writes issued since Bind or ResetWriteCount
    shownSizes As Collection        'the size each hidden position had, keyed on the position
    heldScreenUpdating As Boolean   'ScreenUpdating as it was when the bracket opened
End Type

Private this As TShowHideLayout


'@section Instantiation
'===============================================================================
'@description The factory.

'@label:create
'@sub-title Bind a layout to one worksheet of one layer
'@param sh Worksheet. The sheet the entries live on.
'@param layer Byte. A ShowHideWorksheetLayer value. Which layer the sheet belongs to.
'@param passwords Optional Passwords. The protection keys of the workbook. The
'  layout writes without a protection bracket when it has none.
'@param baseTableName Optional String. The table name with no print_ prefix,
'  which is what the PRINTSTART anchor is named after. Only the printed layer
'  reads it.
'@return ShowHideLayout. The layout, ready to read and write.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.InvalidArgument When the layer is outside the four values.
'@export
Public Function Create(ByVal sh As Worksheet, _
                       ByVal layer As Byte, _
                       Optional ByVal passManager As Passwords, _
                       Optional ByVal baseTableName As String = vbNullString) As ShowHideLayout

    Dim instance As ShowHideLayout

    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A worksheet is needed to show or hide anything on it"
    End If

    If layer < ShowHideLayerHList Or layer > ShowHideLayerCRF Then
        ThrowError ProjectError.InvalidArgument, _
                   "There is no show/hide layer '" & CStr(layer) & "' in a linelist"
    End If

    Set instance = New ShowHideLayout
    instance.Bind sh, layer, passManager, baseTableName
    Set Create = instance
End Function

'@label:bind
'@sub-title Take the four things the layout works with
'@details
'Public because the factory calls it on a fresh instance.
'@param sh Worksheet. The sheet the entries live on.
'@param layer Byte. A ShowHideWorksheetLayer value. The layer of the sheet.
'@param passwords Passwords. The protection keys, or Nothing.
'@param baseTableName String. The table name with no print_ prefix.
'@export
Public Sub Bind(ByVal sh As Worksheet, _
                ByVal layer As Byte, _
                ByVal passManager As Passwords, _
                ByVal baseTableName As String)
    Set this.sh = sh
    this.layer = layer
    this.axisValue = AxisOf(layer)
    Set this.pass = passManager
    this.baseTableName = Trim$(baseTableName)
    this.lockDepth = 0
    this.failures = 0
    this.writes = 0
    Set this.shownSizes = New Collection
End Sub


'@section Accessors
'===============================================================================
'@description What the layout is bound to.

'@label:wksh
'@prop-title The worksheet this layout writes to
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:layer
'@prop-title The layer this layout was built for
'@export
Public Property Get Layer() As Byte
    Layer = this.layer
End Property

'@label:axis
'@prop-title Whether this layer hides rows or columns
'@export
Public Property Get Axis() As Byte
    Axis = this.axisValue
End Property

'@label:supports-orientation
'@prop-title Whether the sheet has a header that can be turned
'@details
'True on a printed sheet alone. A data entry sheet, a VList and a CRF all read
'their labels straight, so turning a header means nothing there.
'@export
Public Property Get SupportsOrientation() As Boolean
    SupportsOrientation = (this.layer = ShowHideLayerPrinted)
End Property

'@label:failure-count
'@prop-title How many writes Excel refused
'@details
'A position the sheet does not have, or a sheet that stays protected, leaves the
'write undone. The count says how many, so a caller can tell a silent run from a
'run that did nothing.
'@export
Public Property Get FailureCount() As Long
    FailureCount = this.failures
End Property

'@label:write-count
'@prop-title How many writes were issued since the bind or the last reset
'@details
'Every SetHidden, SetSize and SetOrientation counts one, refused or not. A
'caller about to save reads it beside ShowHide.HasChanges: nothing written and
'no choice moved means the store already holds what the sheet shows.
'@export
Public Property Get WriteCount() As Long
    WriteCount = this.writes
End Property

'@label:reset-write-count
'@sub-title Start the write count again, once a save has recorded the sheet
'@export
Public Sub ResetWriteCount()
    this.writes = 0
End Sub

'@label:forget-sizes
'@sub-title Drop every kept size
'@details
'For a caller whose sheet another layout instance may have resized in the
'meantime, such as the show/hide session after the layouts form has restored
'a saved layout. The next SizeWhenShown on a hidden position reads the sheet.
'@export
Public Sub ForgetSizes()
    Set this.shownSizes = New Collection
End Sub


'@section Visibility
'===============================================================================
'@description Reading and writing whether a position is hidden.

'@label:is-hidden
'@sub-title Whether a position is hidden on the sheet
'@param position Long. The column or row number.
'@return Boolean. True when Excel is hiding it.
'@export
Public Function IsHidden(ByVal position As Long) As Boolean
    Dim answer As Boolean

    If position < 1 Then Exit Function

    On Error Resume Next
    If this.axisValue = ShowHideAxisRow Then
        answer = this.sh.Rows(position).Hidden
    Else
        answer = this.sh.Columns(position).Hidden
    End If
    On Error GoTo 0

    IsHidden = answer
End Function

'@label:set-hidden
'@sub-title Hide or show one position
'@details
'On the way to hiding, the size the position has is read first and kept on
'the instance, so SizeWhenShown can answer it later without showing the
'column again. Showing drops the kept size, because the sheet answers it.
'@param position Long. The column or row number.
'@param hidden Boolean. True to hide it.
'@export
Public Sub SetHidden(ByVal position As Long, ByVal hidden As Boolean)
    Dim sizeNow As Double

    If position < 1 Then Exit Sub

    If hidden Then
        sizeNow = Size(position)
        If sizeNow > 0 Then KeepShownSize position, sizeNow
    Else
        ForgetShownSize position
    End If

    WriteHidden position, hidden
    this.writes = this.writes + 1
End Sub


'@section Size
'===============================================================================
'@description Reading and writing the width of a column or the height of a row.

'@label:size
'@sub-title The width or height of one position
'@param position Long. The column or row number.
'@return Double. The column width or the row height. A hidden position answers
'  0, which is how Excel reports it.
'@export
Public Function Size(ByVal position As Long) As Double
    Dim answer As Double

    If position < 1 Then Exit Function

    On Error Resume Next
    If this.axisValue = ShowHideAxisRow Then
        answer = this.sh.Rows(position).RowHeight
    Else
        answer = this.sh.Columns(position).ColumnWidth
    End If
    On Error GoTo 0

    Size = answer
End Function

'@label:set-size
'@sub-title Set the width or height of one position
'@details
'A size at or below zero is refused. Excel hides a column that is set to width
'0, so writing a missing size would hide the very entry the caller is trying to
'restore.
'@param position Long. The column or row number.
'@param newSize Double. The width or height to write.
'@export
Public Sub SetSize(ByVal position As Long, ByVal newSize As Double)
    If position < 1 Then Exit Sub
    If newSize <= 0 Then Exit Sub

    UnprotectSheet
    On Error Resume Next
    If this.axisValue = ShowHideAxisRow Then
        this.sh.Rows(position).RowHeight = newSize
    Else
        this.sh.Columns(position).ColumnWidth = newSize
    End If
    If Err.Number <> 0 Then
        this.failures = this.failures + 1
    Else
        'A hidden position answers this size from now on
        KeepShownSize position, newSize
    End If
    On Error GoTo 0
    ReprotectSheet
    this.writes = this.writes + 1
End Sub

'@label:size-when-shown
'@sub-title The size a position would have if it were visible
'@details
'A hidden position reports a size of 0, so a migration file built from a live
'sheet used to record 0 for every hidden column and bring it back at width 0 -
'a column the user can neither see nor grab. A position this instance hid
'answers the size it read on the way to hiding it, with no write. One it found
'hidden already is revealed long enough to read its real size and hidden again;
'those two writes are not counted, since the sheet ends where it began.
'@param position Long. The column or row number.
'@return Double. The size the position holds when visible.
'@export
Public Function SizeWhenShown(ByVal position As Long) As Double
    Dim wasHidden As Boolean
    Dim answer As Double

    If position < 1 Then Exit Function

    wasHidden = IsHidden(position)
    If Not wasHidden Then
        SizeWhenShown = Size(position)
        Exit Function
    End If

    answer = HeldShownSize(position)
    If answer > 0 Then
        SizeWhenShown = answer
        Exit Function
    End If

    BeginBatch
    WriteHidden position, False
    answer = Size(position)
    WriteHidden position, True
    EndBatch

    If answer > 0 Then KeepShownSize position, answer
    SizeWhenShown = answer
End Function


'@section Orientation
'===============================================================================
'@description The turned header of a printed sheet.

'@label:is-vertical
'@sub-title Whether the printed header of one position is turned
'@details
'Excel stores a flat cell two ways: the angle 0, and the constant xlHorizontal
'(-4128), which is what a cell written with `Orientation = 0` reads back as.
'Both mean flat, so both answer False here. Every flat header used to read as
'turned, because the test was `angle <> 0`.
'@param position Long. The column number.
'@return Boolean. True when the header cell reads bottom to top.
'@export
Public Function IsVertical(ByVal position As Long) As Boolean
    Dim headerCell As Range
    Dim angle As Long

    If Not SupportsOrientation Then Exit Function

    Set headerCell = PrintHeaderCell(position)
    If headerCell Is Nothing Then Exit Function

    On Error Resume Next
    angle = CLng(headerCell.Orientation)
    On Error GoTo 0

    IsVertical = (angle <> 0) And (angle <> xlHorizontal)
End Function

'@label:set-orientation
'@sub-title Turn the printed header of one position, or lay it flat
'@details
'The row height and the column autofit belong to the turned branch alone. They
'used to run on every action, so choosing Hide on a column whose header was
'turned laid that header flat on the way to hiding it, and the direction was
'gone when the user showed the column again.
'@param position Long. The column number.
'@param vertical Boolean. True to turn the header.
'@export
Public Sub SetOrientation(ByVal position As Long, ByVal vertical As Boolean)
    Dim headerCell As Range

    If Not SupportsOrientation Then Exit Sub

    Set headerCell = PrintHeaderCell(position)
    If headerCell Is Nothing Then Exit Sub

    UnprotectSheet
    On Error Resume Next
    If vertical Then
        headerCell.Orientation = 90
        headerCell.EntireRow.RowHeight = VERTICAL_HEADER_HEIGHT
        headerCell.EntireColumn.AutoFit
    Else
        headerCell.Orientation = 0
    End If
    If Err.Number <> 0 Then this.failures = this.failures + 1
    On Error GoTo 0
    ReprotectSheet
    this.writes = this.writes + 1
End Sub


'@section Protection bracket
'===============================================================================
'@description Holding the sheet open across a run of writes.

'@label:begin-batch
'@sub-title Unprotect the sheet and keep it open
'@details
'A loop that writes many positions calls this first, so the sheet is unprotected
'once instead of once per position. Every BeginBatch needs its EndBatch.
'@export
Public Sub BeginBatch()
    UnprotectSheet
End Sub

'@label:end-batch
'@sub-title Close the bracket BeginBatch opened
'@export
Public Sub EndBatch()
    ReprotectSheet
End Sub


'@section Helpers
'===============================================================================
'@description Private workers.

'@label:write-hidden
'@sub-title The one write of the hidden state, inside its bracket
'@details
'SetHidden and SizeWhenShown both come through here. The size bookkeeping and
'the write count stay in SetHidden, so the show-and-hide pair SizeWhenShown
'issues to measure a column counts nothing and keeps nothing.
'@param position Long. The column or row number.
'@param hidden Boolean. True to hide it.
Private Sub WriteHidden(ByVal position As Long, ByVal hidden As Boolean)
    UnprotectSheet
    On Error Resume Next
    If this.axisValue = ShowHideAxisRow Then
        this.sh.Rows(position).Hidden = hidden
    Else
        this.sh.Columns(position).Hidden = hidden
    End If
    If Err.Number <> 0 Then this.failures = this.failures + 1
    On Error GoTo 0
    ReprotectSheet
End Sub

'@label:keep-shown-size
'@sub-title Record the size one position had before it was hidden
'@param position Long. The column or row number.
'@param shownSize Double. The size read while it was visible.
Private Sub KeepShownSize(ByVal position As Long, ByVal shownSize As Double)
    ForgetShownSize position
    this.shownSizes.Add shownSize, CStr(position)
End Sub

'@label:forget-shown-size
'@sub-title Drop the kept size of one position, if any
'@param position Long. The column or row number.
Private Sub ForgetShownSize(ByVal position As Long)
    On Error Resume Next
    this.shownSizes.Remove CStr(position)
    On Error GoTo 0
End Sub

'@label:held-shown-size
'@sub-title The kept size of one position, or 0 when none was kept
'@param position Long. The column or row number.
'@return Double. The size read before the hide, or 0.
Private Function HeldShownSize(ByVal position As Long) As Double
    On Error Resume Next
    HeldShownSize = this.shownSizes.Item(CStr(position))
    On Error GoTo 0
End Function

'@label:axis-of
'@sub-title The direction one layer occupies
'@param layer Byte. A ShowHideWorksheetLayer value. The layer to answer for.
'@return Byte. A ShowHideAxis value. Column for HList and Printed, row for VList and CRF.
Private Function AxisOf(ByVal layer As Byte) As Byte
    Select Case layer
        Case ShowHideLayerVList, ShowHideLayerCRF
            AxisOf = ShowHideAxisRow
        Case Else
            AxisOf = ShowHideAxisColumn
    End Select
End Function

'@label:print-header-cell
'@sub-title The printed header cell of one column
'@details
'The anchor is a named cell called `<table name>_PRINTSTART`, written by
'VarWriter on the first line of the printed table, and the table name it uses
'carries no print_ prefix. The header the reader sees is the line above it.
'@param position Long. The column number.
'@return Range. The header cell, or Nothing when the anchor is absent.
Private Function PrintHeaderCell(ByVal position As Long) As Range
    Dim anchor As Range

    If position < 1 Then Exit Function
    If LenB(this.baseTableName) = 0 Then Exit Function

    On Error Resume Next
    Set anchor = this.sh.Range(this.baseTableName & PRINT_ANCHOR_SUFFIX)
    On Error GoTo 0

    If anchor Is Nothing Then Exit Function

    On Error Resume Next
    Set PrintHeaderCell = this.sh.Cells(anchor.Offset(PRINT_HEADER_OFFSET).Row, position)
    On Error GoTo 0
End Function

'@label:unprotect-sheet
'@sub-title Unprotect the sheet, counting how deep the bracket is
'@details
'Not called Unlock. Lock and Unlock are VBA file statements, so a line holding
'nothing but the word Unlock is read as one of those and the compile breaks.
'
'THE SCREEN IS HELD HERE TOO
'-------------------------------------------------------------------------------
'Hiding a column repaints the sheet. Every write in this class comes through
'this bracket, so the screen is held from the moment the sheet opens until the
'moment it closes again, and a run of writes repaints once at the end instead of
'once per position. That is what made the show/hide forms flicker: applying a
'saved layout writes every entry in turn, and SizeWhenShown deliberately shows a
'column and hides it again to measure it, which was a visible flash per column.
'
'The value the bracket opened on is what gets put back, so a caller that had
'already turned the screen off keeps it off when the bracket closes.
Private Sub UnprotectSheet()
    this.lockDepth = this.lockDepth + 1
    If this.lockDepth > 1 Then Exit Sub

    On Error Resume Next
    this.heldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo 0

    If this.pass Is Nothing Then Exit Sub

    On Error Resume Next
    this.pass.UnProtect this.sh.Name
    On Error GoTo 0
End Sub

'@label:reprotect-sheet
'@sub-title Protect the sheet and release the screen once the outermost bracket closes
Private Sub ReprotectSheet()
    this.lockDepth = this.lockDepth - 1
    If this.lockDepth > 0 Then Exit Sub

    this.lockDepth = 0

    If Not this.pass Is Nothing Then
        On Error Resume Next
        this.pass.Protect this.sh.Name
        On Error GoTo 0
    End If

    'Last, so the one repaint the user sees happens with the sheet already
    'protected and every position already written.
    On Error Resume Next
    Application.ScreenUpdating = this.heldScreenUpdating
    On Error GoTo 0
End Sub

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
