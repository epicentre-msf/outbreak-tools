VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Hiding and showing whole sections of one worksheet"

'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Hiding and showing whole sections of one worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit


'@class SectionShowHide
'@description
'One section of one worksheet, hidden or shown. The three pieces it needs are
'already classes: SectionMap says where each section starts and ends, ShowHide
'holds the choice for every variable, and ShowHideLayout writes the sheet. This
'binds the three so the section button and the section form share one body.
'
'WHERE THE USER HAS TO STAND
'-------------------------------------------------------------------------------
'A section title is written on one line of the sheet and nowhere else:
'
'  HList   row 5
'  VList   column 2
'
'SectionBuilder writes them there and the two numbers are copied here from it.
'SectionAtCell answers a block only for a cell on that line, so a click made
'from a data row changes nothing.
'
'WHY THE STATE IS READ AND NEVER STORED
'-------------------------------------------------------------------------------
'Hiding a section stores nothing of its own. The choice lands on each variable
'inside it, and State reads those variables back. So a section a user collapsed
'through the form, through the button, or by hiding its columns one at a time
'all read the same, and nothing can drift out of step.
'
'A section holding no variable the user owns answers Fixed, and CanChange is
'False for it. Offering an action that does nothing is worse than refusing it.
'@depends SectionMap, ShowHide, ShowHideLayout
'@version 1.0 (2026-08-14)


Private Const CLASS_NAME As String = "SectionShowHide"

'The line the section titles are written on. Both numbers are SectionBuilder's
'(HLIST_SEC_ROW and VLIST_SEC_COL) and mean nothing without it.
Private Const HLIST_TITLE_ROW As Long = 5
Private Const VLIST_TITLE_COLUMN As Long = 2

Private Type TSectionShowHide
    secMap As SectionMap
    entries As ShowHide
    layout As ShowHideLayout
End Type

Private this As TSectionShowHide


'@section Instantiation
'===============================================================================
'@description The factory.

'@label:create
'@sub-title Bind a section toggle to one worksheet
'@param secMap SectionMap. The sections of the sheet.
'@param entries ShowHide. The variables of the sheet and their choices.
'@param layout ShowHideLayout. The sheet the choices are written to.
'@return SectionShowHide. The toggle, ready to read and write.
'@throws ProjectError.ObjectNotInitialized When any of the three is Nothing.
'@export
Public Function Create(ByVal secMap As SectionMap, _
                       ByVal entries As ShowHide, _
                       ByVal layout As ShowHideLayout) As SectionShowHide

    Dim instance As SectionShowHide

    If secMap Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A section map is needed to hide a section"
    End If

    If entries Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "An entry list is needed to hide a section"
    End If

    If layout Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A worksheet layout is needed to hide a section"
    End If

    Set instance = New SectionShowHide
    instance.Bind secMap, entries, layout
    Set Create = instance
End Function

'@label:bind
'@sub-title Take the three pieces the toggle works with
'@details
'Public because the factory calls it on a fresh instance.
'@param secMap SectionMap. The sections of the sheet.
'@param entries ShowHide. The variables of the sheet and their choices.
'@param layout ShowHideLayout. The sheet the choices are written to.
'@export
Public Sub Bind(ByVal secMap As SectionMap, _
                ByVal entries As ShowHide, _
                ByVal layout As ShowHideLayout)
    Set this.secMap = secMap
    Set this.entries = entries
    Set this.layout = layout
End Sub


'@section Accessors
'===============================================================================
'@description What the toggle is bound to.

'@label:count
'@prop-title How many sections the sheet carries
'@export
Public Property Get Count() As Long
    Count = this.secMap.Count
End Property

'@label:wksh
'@prop-title The worksheet the toggle writes to
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.layout.Wksh
End Property

'@label:section-name-at
'@sub-title The title of one section
'@param index Long. A one based section index.
'@return String. The title SectionBuilder wrote.
'@export
Public Function SectionNameAt(ByVal index As Long) As String
    SectionNameAt = this.secMap.SectionNameAt(index)
End Function


'@section Where the user is standing
'===============================================================================
'@description Turning a selected cell into a section.

'@label:title-line
'@prop-title The row or column the section titles are written on
'@details
'A row number on an HList sheet and a column number on a VList one, which is
'the same split ShowHideLayout makes for the hiding axis itself.
'@return Long. The line the titles sit on.
'@export
Public Property Get TitleLine() As Long
    If this.layout.Axis = ShowHideAxisRow Then
        TitleLine = VLIST_TITLE_COLUMN
    Else
        TitleLine = HLIST_TITLE_ROW
    End If
End Property

'@label:section-at-cell
'@sub-title The section one selected cell stands on
'@details
'The answer is a section only when the cell sits on the title line. A cell in a
'data row belongs to a section too, and reading it as one is what let a second
'press of the button collapse a section the user never pointed at: hiding the
'columns of a section moves the cursor onto the next visible column, and that
'column is inside the next section.
'
'A multi-cell selection is read from its first cell.
'@param cellRng Range. What the user has selected.
'@return Long. A one based section index, or 0 when the cell names no section.
'@export
Public Function SectionAtCell(ByVal cellRng As Range) As Long
    Dim anchor As Range
    Dim position As Long

    If cellRng Is Nothing Then Exit Function

    Set anchor = cellRng.Cells(1, 1)

    If this.layout.Axis = ShowHideAxisRow Then
        If anchor.Column <> VLIST_TITLE_COLUMN Then Exit Function
        position = anchor.Row
    Else
        If anchor.Row <> HLIST_TITLE_ROW Then Exit Function
        position = anchor.Column
    End If

    SectionAtCell = this.secMap.IndexAtPosition(position)
End Function

'@label:title-cell
'@sub-title The cell holding the title of one section
'@details
'Where the caller puts the cursor once a section is shown again, so the next
'press of the button acts on the section the user is looking at.
'@param index Long. A one based section index.
'@return Range. The title cell, or Nothing when the index names no section.
'@export
Public Function TitleCell(ByVal index As Long) As Range
    If index < 1 Or index > this.secMap.Count Then Exit Function

    On Error Resume Next
    If this.layout.Axis = ShowHideAxisRow Then
        Set TitleCell = Wksh().Cells(this.secMap.StartAt(index), VLIST_TITLE_COLUMN)
    Else
        Set TitleCell = Wksh().Cells(HLIST_TITLE_ROW, this.secMap.StartAt(index))
    End If
    On Error GoTo 0
End Function


'@section Reading and writing one section
'===============================================================================
'@description The state of a section, and changing it.

'@label:state
'@sub-title Where one section stands
'@param index Long. A one based section index.
'@return Byte. A ShowHideRangeState value.
'@export
Public Function State(ByVal index As Long) As Byte
    If index < 1 Or index > this.secMap.Count Then
        State = ShowHideRangeEmpty
        Exit Function
    End If

    State = this.entries.RangeState(this.secMap.StartAt(index), _
                                    this.secMap.EndAt(index))
End Function

'@label:is-hidden
'@sub-title Whether one section is hidden right now
'@details
'A section is hidden when every variable the user owns inside it is hidden. A
'section showing some of its variables answers False, so the next action offered
'for it is to hide the rest.
'@param index Long. A one based section index.
'@return Boolean. True when the section is collapsed.
'@export
Public Function IsHidden(ByVal index As Long) As Boolean
    IsHidden = (State(index) = ShowHideRangeHidden)
End Function

'@label:can-change
'@sub-title Whether the user may hide or show one section
'@details
'False for a section holding no variable the user owns, and for an index outside
'the map. Both are sections the toggle would leave exactly as it found them.
'@param index Long. A one based section index.
'@return Boolean. True when the section follows the user.
'@export
Public Function CanChange(ByVal index As Long) As Boolean
    Dim sectionState As Byte

    sectionState = State(index)
    CanChange = (sectionState <> ShowHideRangeEmpty) And _
                (sectionState <> ShowHideRangeFixed)
End Function

'@label:set-hidden
'@sub-title Hide or show one whole section
'@details
'The choice lands on every variable of the section the user owns, and the sheet
'is then put in step through the layout. A mandatory or a locked variable keeps
'its state, so a section holding one never disappears whole.
'@param index Long. A one based section index.
'@param hidden Boolean. True to collapse the section.
'@return Long. How many variables the call changed.
'@export
Public Function SetHidden(ByVal index As Long, ByVal hidden As Boolean) As Long
    Dim changed As Long

    If index < 1 Or index > this.secMap.Count Then Exit Function

    changed = this.entries.SetHiddenInRange(this.secMap.StartAt(index), _
                                            this.secMap.EndAt(index), _
                                            hidden)
    this.entries.Apply this.layout

    SetHidden = changed
End Function

'@label:set-all-hidden
'@sub-title Hide or show every section of the sheet
'@details
'One pass over the map and one pass over the sheet. The section form's two
'option buttons reach it when no single section is picked.
'@param hidden Boolean. True to collapse them all.
'@return Long. How many variables the call changed.
'@export
Public Function SetAllHidden(ByVal hidden As Boolean) As Long
    Dim index As Long
    Dim changed As Long

    For index = 1 To this.secMap.Count
        changed = changed + _
                  this.entries.SetHiddenInRange(this.secMap.StartAt(index), _
                                                this.secMap.EndAt(index), _
                                                hidden)
    Next index

    this.entries.Apply this.layout

    SetAllHidden = changed
End Function


'@section Helpers
'===============================================================================
'@description Private workers.

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
