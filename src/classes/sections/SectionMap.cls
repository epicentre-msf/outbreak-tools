VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Where each section of one worksheet begins and ends"

'@PredeclaredId
'@Folder("Sections")
'@ModuleDescription("Where each section of one worksheet begins and ends")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit

'@class SectionMap
'@description
'The first and last position of every section of one worksheet, written when
'the sheet is built and read back whenever something needs to act on a whole
'section. Add records one block, Count, SectionNameAt, StartAt and EndAt read
'one back, IndexAtPosition finds the block a position falls in, and
'IndicesInRange finds every block a span touches. Clear drops the lot.
'
'A SECTION IS A BLOCK OF POSITIONS
'-------------------------------------------------------------------------------
'A block is identified by its position range, and the title is carried for a
'person to read. Two reasons for that way round:
'
'The caller holds a cell address. The section toggle starts from what the user
'has selected, so the question it asks is "which section covers this column",
'and a position is what it has to ask with.
'
'SectionBuilder groups CONSECUTIVE dictionary rows, so a run of rows is what a
'block means to the class that writes it. `LLdictionary.Prepare` sorts the
'dictionary by sheet and main section before it derives `column index`, which
'gives one run per title today. Keying on the range holds either way and costs
'the same.
'
'WHAT A POSITION MEANS
'-------------------------------------------------------------------------------
'The same thing it means to ShowHide and ShowHideLayout: a column number on an
'HList sheet and a row number on a VList one. Both come from the dictionary
'`column index` column, which is what SectionBuilder lays its headers out on.
'
'WHERE THE MAP LIVES
'-------------------------------------------------------------------------------
'In the hidden names of the sheet itself, one name per block:
'
'  section_map_count   how many blocks the sheet carries
'  section_map_<n>     "<start>|<end>|<name>" for block n
'
'One name carries one whole block. A defined name holds its value in its
'RefersTo formula and that string is length capped, so a sheet with thirty
'sections would run a single name holding them all past the cap. The section
'title is written last and truncated, which keeps the two numbers safe even
'when a designer writes a long one.
'@depends HiddenNames, BetterArray, Checking
'@version 1.0 (2026-08-06)


Private Const CLASS_NAME As String = "SectionMap"

'The hidden names the map is kept under
Private Const COUNT_NAME As String = "section_map_count"
Private Const BLOCK_PREFIX As String = "section_map_"

'What separates the two numbers and the name inside one stored block
Private Const FIELD_SEPARATOR As String = "|"

'How many fields a stored block holds. The title is the last field, so the
'split is capped here and a title carrying the separator arrives whole.
Private Const FIELD_COUNT As Long = 3

'How much of a section name is kept. A defined name holds its value in its
'RefersTo formula and that string is length capped, so the name is cut before
'it can push the two numbers out.
Private Const MAX_NAME_LENGTH As Long = 180

Private Type TSectionMap
    sh As Worksheet
    hn As HiddenNames
    starts() As Long
    ends() As Long
    names() As String
    count As Long
End Type

Private this As TSectionMap


'@section Instantiation
'===============================================================================
'@description The factory and the load it runs.

'@label:create
'@sub-title Read the section map of one worksheet
'@details
'A sheet with no map on it answers a Count of 0. Every reader tests the count
'before it reads a block, so an unbuilt sheet and a sheet built before this
'class existed both come back through the same path.
'@param wksh Worksheet. The sheet the map belongs to.
'@return SectionMap. The map, ready to read and to add to.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@export
Public Function Create(ByVal wksh As Worksheet) As SectionMap
    Dim instance As SectionMap

    If wksh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A worksheet is needed to read its section map"
    End If

    Set instance = New SectionMap
    instance.Initialise wksh
    Set Create = instance
End Function

'@label:initialise
'@sub-title Bind to a sheet and load the blocks it carries
'@details
'Public because the factory calls it on a fresh instance.
'@param wksh Worksheet. The sheet the map belongs to.
'@export
Public Sub Initialise(ByVal wksh As Worksheet)
    Set this.sh = wksh
    Set this.hn = HiddenNames.Create(wksh)
    LoadBlocks
End Sub


'@section Accessors
'===============================================================================
'@description What the map holds.

'@label:wksh
'@prop-title The worksheet this map belongs to
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:count
'@prop-title How many section blocks the sheet carries
'@export
Public Property Get Count() As Long
    Count = this.count
End Property

'@label:section-name-at
'@sub-title The title of one block
'@param index Long. A one based block index.
'@return String. The `main section` value the block was built from. A section
'  the dictionary left blank answers an empty string.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function SectionNameAt(ByVal index As Long) As String
    EnsureValidIndex index
    SectionNameAt = this.names(index)
End Function

'@label:start-at
'@sub-title The first position of one block
'@param index Long. A one based block index.
'@return Long. The column number on an HList sheet, the row number on a VList one.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function StartAt(ByVal index As Long) As Long
    EnsureValidIndex index
    StartAt = this.starts(index)
End Function

'@label:end-at
'@sub-title The last position of one block
'@param index Long. A one based block index.
'@return Long. The column number on an HList sheet, the row number on a VList one.
'@throws ProjectError.InvalidArgument When the index is out of range.
'@export
Public Function EndAt(ByVal index As Long) As Long
    EnsureValidIndex index
    EndAt = this.ends(index)
End Function


'@section Lookup
'===============================================================================
'@description Finding the block a selection lands in.

'@label:index-at-position
'@sub-title The block one position falls in
'@param position Long. A column number or a row number.
'@return Long. The one based block index, or 0 when no block covers the position.
'@export
Public Function IndexAtPosition(ByVal position As Long) As Long
    Dim idx As Long

    If position < 1 Then Exit Function

    For idx = 1 To this.count
        If position >= this.starts(idx) And position <= this.ends(idx) Then
            IndexAtPosition = idx
            Exit Function
        End If
    Next idx
End Function

'@label:indices-in-range
'@sub-title Every block a span of positions touches
'@details
'What a selection of more than one cell asks. A block is answered when it
'overlaps the span at all, which is what lets a user reach a section whose
'every position is hidden: selecting from the visible position before it to the
'visible position after it covers the whole gap.
'
'The two bounds are read in either order, so a selection made right to left
'answers the same blocks as one made left to right.
'@param firstPos Long. One end of the span.
'@param lastPos Long. The other end of the span.
'@return BetterArray. The one based block indices, in map order. Empty when the
'  span touches no block.
'@export
Public Function IndicesInRange(ByVal firstPos As Long, ByVal lastPos As Long) As BetterArray
    Dim found As BetterArray
    Dim idx As Long
    Dim lowPos As Long
    Dim highPos As Long

    Set found = New BetterArray
    found.LowerBound = 1
    Set IndicesInRange = found

    lowPos = firstPos
    highPos = lastPos
    If lowPos > highPos Then
        lowPos = lastPos
        highPos = firstPos
    End If

    If highPos < 1 Then Exit Function

    For idx = 1 To this.count
        If (this.starts(idx) <= highPos) And (this.ends(idx) >= lowPos) Then
            found.Push idx
        End If
    Next idx
End Function


'@section Writing
'===============================================================================
'@description What SectionBuilder calls while it builds a sheet.

'@label:clear
'@sub-title Drop every block the sheet carries
'@details
'Called before a rebuild. A sheet rebuilt with fewer sections than last time
'would otherwise keep the tail of the old map, and those stale blocks point at
'positions the new sheet gives to something else.
'@export
Public Sub Clear()
    Dim idx As Long

    For idx = 1 To this.count
        RemoveIfPresent BlockName(idx)
    Next idx

    this.count = 0
    Erase this.starts
    Erase this.ends
    Erase this.names

    WriteCount 0
End Sub

'@label:add
'@sub-title Record one section block
'@param sectionName String. The `main section` value, or an empty string.
'@param startPos Long. The first position of the block.
'@param endPos Long. The last position of the block.
'@return Long. The one based index the block was filed under.
'@throws ProjectError.InvalidArgument When startPos is below 1 or endPos is
'  before startPos.
'@export
Public Function Add(ByVal sectionName As String, _
                    ByVal startPos As Long, _
                    ByVal endPos As Long) As Long

    Dim idx As Long

    If startPos < 1 Then
        ThrowError ProjectError.InvalidArgument, _
                   "A section starts at position 1 or later, not " & CStr(startPos)
    End If

    If endPos < startPos Then
        ThrowError ProjectError.InvalidArgument, _
                   "A section ending at " & CStr(endPos) & _
                   " cannot start at " & CStr(startPos)
    End If

    idx = this.count + 1
    ReDim Preserve this.starts(1 To idx)
    ReDim Preserve this.ends(1 To idx)
    ReDim Preserve this.names(1 To idx)

    this.starts(idx) = startPos
    this.ends(idx) = endPos
    this.names(idx) = TrimmedName(sectionName)
    this.count = idx

    WriteValue BlockName(idx), Encode(this.names(idx), startPos, endPos)
    WriteCount idx

    Add = idx
End Function


'@section Storage
'===============================================================================
'@description Reading and writing the hidden names the map is kept in.

'@label:load-blocks
'@sub-title Read every stored block into memory
'@details
'A block whose name is missing or unreadable ends the load. The count and the
'blocks are written together, so a gap means the sheet was interrupted
'mid-build and what follows the gap cannot be trusted.
Private Sub LoadBlocks()
    Dim stored As Long
    Dim idx As Long
    Dim raw As String
    Dim startPos As Long
    Dim endPos As Long
    Dim sectionName As String

    this.count = 0
    Erase this.starts
    Erase this.ends
    Erase this.names

    stored = ReadCount()
    If stored < 1 Then Exit Sub

    ReDim this.starts(1 To stored)
    ReDim this.ends(1 To stored)
    ReDim this.names(1 To stored)

    For idx = 1 To stored
        raw = ReadValue(BlockName(idx))
        If Not Decode(raw, sectionName, startPos, endPos) Then Exit For

        this.starts(idx) = startPos
        this.ends(idx) = endPos
        this.names(idx) = sectionName
        this.count = idx
    Next idx

    If this.count = 0 Then
        Erase this.starts
        Erase this.ends
        Erase this.names
    ElseIf this.count < stored Then
        ReDim Preserve this.starts(1 To this.count)
        ReDim Preserve this.ends(1 To this.count)
        ReDim Preserve this.names(1 To this.count)
    End If
End Sub

'@label:block-name
'@sub-title The hidden name one block is kept under
'@param index Long. A one based block index.
'@return String. The name identifier.
Private Function BlockName(ByVal index As Long) As String
    BlockName = BLOCK_PREFIX & CStr(index)
End Function

'@label:encode
'@sub-title Write a block as one stored string
'@param sectionName String. The section title, already trimmed to length.
'@param startPos Long. The first position.
'@param endPos Long. The last position.
'@return String. The stored form.
Private Function Encode(ByVal sectionName As String, _
                        ByVal startPos As Long, _
                        ByVal endPos As Long) As String
    Encode = CStr(startPos) & FIELD_SEPARATOR & CStr(endPos) & _
             FIELD_SEPARATOR & sectionName
End Function

'@label:decode
'@sub-title Read a stored string back into a block
'@details
'The split is capped at three fields, which lets the last field hold a section
'title carrying the separator and still arrive whole.
'@param raw String. The stored form.
'@param sectionName String. Set to the section title.
'@param startPos Long. Set to the first position.
'@param endPos Long. Set to the last position.
'@return Boolean. True when the string held a usable block.
Private Function Decode(ByVal raw As String, _
                        ByRef sectionName As String, _
                        ByRef startPos As Long, _
                        ByRef endPos As Long) As Boolean

    Dim parts() As String

    sectionName = vbNullString
    startPos = 0
    endPos = 0

    If LenB(raw) = 0 Then Exit Function

    parts = Split(raw, FIELD_SEPARATOR, FIELD_COUNT)
    If UBound(parts) < 1 Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    startPos = CLng(parts(0))
    endPos = CLng(parts(1))
    If startPos < 1 Or endPos < startPos Then
        startPos = 0
        endPos = 0
        Exit Function
    End If

    If UBound(parts) >= 2 Then sectionName = parts(2)

    Decode = True
End Function

'@label:trimmed-name
'@sub-title Cut a section title down to what a defined name can carry
'@param sectionName String. The title as the dictionary spells it.
'@return String. The title, at most MAX_NAME_LENGTH characters.
Private Function TrimmedName(ByVal sectionName As String) As String
    Dim answer As String

    answer = Trim$(sectionName)
    If Len(answer) > MAX_NAME_LENGTH Then answer = Left$(answer, MAX_NAME_LENGTH)

    TrimmedName = answer
End Function

'@label:read-count
'@sub-title Read how many blocks the sheet says it carries
'@return Long. The stored count, or 0.
Private Function ReadCount() As Long
    If this.hn Is Nothing Then Exit Function
    ReadCount = this.hn.ValueAsLong(COUNT_NAME, 0)
End Function

'@label:write-count
'@sub-title Write how many blocks the sheet carries
'@param newCount Long. The count to store.
Private Sub WriteCount(ByVal newCount As Long)
    If this.hn Is Nothing Then Exit Sub
    this.hn.EnsureName COUNT_NAME, newCount, HiddenNameTypeLong
    this.hn.SetValue COUNT_NAME, newCount
End Sub

'@label:read-value
'@sub-title Read one stored block, answering empty when it is absent
'@param nameId String. The hidden name to read.
'@return String. The stored form, or an empty string.
Private Function ReadValue(ByVal nameId As String) As String
    If this.hn Is Nothing Then Exit Function
    ReadValue = this.hn.ValueAsString(nameId, vbNullString)
End Function

'@label:write-value
'@sub-title Write one stored block
'@details
'EnsureName then SetValue, because EnsureName leaves a name that already exists
'holding whatever it held before. Clear removes the old names first, so this is
'the belt to that pair of braces.
'@param nameId String. The hidden name to write.
'@param value String. The stored form.
Private Sub WriteValue(ByVal nameId As String, ByVal value As String)
    If this.hn Is Nothing Then Exit Sub
    this.hn.EnsureName nameId, value, HiddenNameTypeString
    this.hn.SetValue nameId, value
End Sub

'@label:remove-if-present
'@sub-title Drop one stored block when the sheet carries it
'@param nameId String. The hidden name to drop.
Private Sub RemoveIfPresent(ByVal nameId As String)
    If this.hn Is Nothing Then Exit Sub
    If Not this.hn.HasName(nameId) Then Exit Sub
    this.hn.RemoveName nameId
End Sub


'@section Helpers
'===============================================================================
'@description Private guards.

'@label:ensure-valid-index
'@sub-title Refuse a block index outside the map
'@param index Long. The index to check.
'@throws ProjectError.InvalidArgument When the index is out of range.
Private Sub EnsureValidIndex(ByVal index As Long)
    If index < 1 Or index > this.count Then
        ThrowError ProjectError.InvalidArgument, _
                   "Section index " & CStr(index) & " is outside range 1.." & CStr(this.count)
    End If
End Sub

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
