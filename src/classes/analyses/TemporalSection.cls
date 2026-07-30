VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemporalSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Holds the date bounds and the headers of one temporal analysis section"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Holds the date bounds and the headers of one temporal analysis section")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName

Option Explicit

'@class TemporalSection
'@description
'One section of a time series or spatio-temporal analysis sheet. A section is a
'run of tables that share a date period, and it carries three things: the two
'formulas that give the period its bounds, the list of table headers its
'navigation dropdown offers, and the identifiers of its first and last table.
'
'THE SECTION IS STATED, NOT GUESSED
'-------------------------------------------------------------------------------
'AnalysisOutput used to work out where a section ended by building the
'specification of the row after the one it was on and asking whether that row
'started a new section. Five defects came out of that one look-ahead: the flush
'branch tested a condition that was always False, the header list was never
'cleared so every dropdown listed every header of the sheet, the last table of
'the last section was left out of the date bounds, a section holding one table
'pushed its header twice, and the previous section identifier was carried
'between iterations and never reset.
'
'The loop says what it knows instead. It sees a new section directly, so it
'closes the one it holds and starts the next. Every table it draws is added
'exactly once. Nothing looks ahead and nothing walks backwards, which is also
'what removes the two quadratic specification walks the loop paid for.
'
'NO WORKSHEET REACHES THIS CLASS
'-------------------------------------------------------------------------------
'This holds strings and a list. The caller reads MinFormula, MaxFormula and
'Headers when it closes a section, and writes them where they belong. That is
'what lets the whole state machine be tested with no workbook, no linelist and
'no dictionary.
'@depends BetterArray
'@version 2026-07-30

'@section State
'===============================================================================

Private Type TTemporalSection
    sectionId As String
    lastTableId As String
    minParts As String
    maxParts As String
    tableCount As Long
    headers As BetterArray
    sectionOpen As Boolean
End Type

Private this As TTemporalSection

Private Const CLASS_NAME As String = "TemporalSection"

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a holder with no section started
'@details
'StartSection begins the first section.
'@return TemporalSection. A holder holding nothing.
'@export
Public Function Create() As TemporalSection
    Dim instance As TemporalSection
    Set instance = New TemporalSection
    Set Create = instance
End Function

'@section Section lifecycle
'===============================================================================

'@label:start-section
'@sub-title Begin a section on its first table
'@details
'Clears the bounds and the headers of the section that came before, so nothing
'of it can reach the next dropdown. The identifier is the table identifier of
'the row that starts the section, which is what names every range the section
'owns on the sheet.
'@param secId String. Table identifier of the first table of the section.
'@throws ProjectError.ErrorUnexpectedState When a section is already open.
'@throws ProjectError.InvalidArgument When the identifier is empty.
'@export
Public Sub StartSection(ByVal secId As String)
    If this.sectionOpen Then _
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "Section " & this.sectionId & " is still open"
    If LenB(secId) = 0 Then _
        ThrowError ProjectError.InvalidArgument, "A section needs an identifier"

    this.sectionId = secId
    this.lastTableId = vbNullString
    this.minParts = vbNullString
    this.maxParts = vbNullString
    this.tableCount = 0
    Set this.headers = New BetterArray
    this.headers.LowerBound = 1
    this.sectionOpen = True
End Sub

'@label:add-table
'@sub-title Add one table to the open section
'@details
'Appends the table's date column to both bounds and pushes its header onto the
'dropdown list. Every table of the section comes through here exactly once, the
'first and the last included.
'@param tabId String. Table identifier of the table being added.
'@param tableName String. Name of the linelist table the dates come from.
'@param rowVar String. Name of the date column inside that table.
'@param header String. Display string for the navigation dropdown.
'@throws ProjectError.ErrorUnexpectedState When no section is open.
'@export
Public Sub AddTable(ByVal tabId As String, ByVal tableName As String, _
                    ByVal rowVar As String, ByVal header As String)
    Dim dateColumn As String

    If Not this.sectionOpen Then _
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "No section is open to add a table to"

    dateColumn = tableName & "[" & rowVar & "]"

    If this.tableCount = 0 Then
        this.minParts = "MIN(" & dateColumn & ")"
        this.maxParts = "MAX(" & dateColumn & ")"
    Else
        this.minParts = this.minParts & ", MIN(" & dateColumn & ")"
        this.maxParts = this.maxParts & ", MAX(" & dateColumn & ")"
    End If

    this.tableCount = this.tableCount + 1
    this.lastTableId = tabId
    this.headers.Push header
End Sub

'@label:end-section
'@sub-title End the open section
'@details
'The bounds, the headers and the two identifiers stay readable after this, so
'the caller ends the section and then writes what it held. Starting the next
'section is what clears them.
'@throws ProjectError.ErrorUnexpectedState When no section is open.
'@export
Public Sub EndSection()
    If Not this.sectionOpen Then _
        ThrowError ProjectError.ErrorUnexpectedState, "No section is open to end"
    this.sectionOpen = False
End Sub

'@section Section values
'===============================================================================

'@label:is-open
'@prop-title Whether a section is open
'@return Boolean. True between StartSection and EndSection.
'@export
Public Property Get IsOpen() As Boolean
    IsOpen = this.sectionOpen
End Property

'@label:section-id
'@prop-title Identifier of the section
'@return String. Table identifier of the first table of the section.
'@export
Public Property Get SectionId() As String
    SectionId = this.sectionId
End Property

'@label:last-table-id
'@prop-title Identifier of the last table added to the section
'@return String. Empty until a table has been added.
'@export
Public Property Get LastTableId() As String
    LastTableId = this.lastTableId
End Property

'@label:table-count
'@prop-title How many tables the section holds
'@return Long. Zero on a section with no table.
'@export
Public Property Get TableCount() As Long
    TableCount = this.tableCount
End Property

'@label:min-formula
'@prop-title Worksheet formula for the earliest date of the section
'@details
'Answers an empty string on a section holding no table, so a caller can tell
'"nothing to write" from a formula.
'@return String. The complete formula, or an empty string.
'@export
Public Property Get MinFormula() As String
    If this.tableCount = 0 Then Exit Property
    MinFormula = "= MIN(" & this.minParts & ")"
End Property

'@label:max-formula
'@prop-title Worksheet formula for the latest date of the section
'@return String. The complete formula, or an empty string.
'@export
Public Property Get MaxFormula() As String
    If this.tableCount = 0 Then Exit Property
    MaxFormula = "= MAX(" & this.maxParts & ")"
End Property

'@label:headers
'@prop-title Dropdown entries of the section
'@details
'A copy, so a caller holding on to it cannot grow the section's own list.
'@return BetterArray. One entry per table, in the order the tables were added.
'@export
Public Property Get Headers() As BetterArray
    Dim entries As BetterArray

    If this.headers Is Nothing Then
        Set entries = New BetterArray
        entries.LowerBound = 1
        Set Headers = entries
        Exit Property
    End If

    Set Headers = this.headers.Clone()
End Property

'@section Error handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber Long. The error code to raise.
'@param message String. What went wrong.
'@throws ProjectError.<varies> Always raises.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
