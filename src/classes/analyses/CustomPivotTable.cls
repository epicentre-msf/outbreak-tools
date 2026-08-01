VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomPivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Custom pivot table creation on a dedicated worksheet"

'@PredeclaredId
'@Folder("Tables")
'@ModuleDescription("Custom pivot table creation on a dedicated worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class CustomPivotTable
'@description
'Manages pivot table creation on a dedicated worksheet during linelist
'building. Each call to Add creates a new PivotTable block (title + pivot)
'stacked vertically, and decides whether the sheet has earned its place in
'front of the user.
'
'PUBLISHED STATE
'-------------------------------------------------------------------------------
'Every fact this class records lives in a worksheet-level HiddenName on the
'pivot sheet. Nothing is written into the sheet's cells except the pivot
'tables themselves and their titles. Three of the five names are read from
'outside this class, so renaming one is a breaking change:
'
'  pivot_output_row    Long     private   the next block's base row
'  pivot_counter       Long     private   the number printed in the next title
'  pivot_has_content   Boolean  private   set once a pivot has actually landed
'  pivot_formatted     Boolean  private   guards the once-only cosmetic pass
'  pivot_title_<table> String   PUBLIC    read by EventsLinelistButtons
'                                         (ClickOpenVarLab), which lists every
'                                         variable under its pivot's title
'
'An earlier version of this class kept the first two in cell A1 and named the
'title cell as a range. Both were read from outside and both broke when the
'state moved. The table above is the contract; keep it accurate.
'
'@depends LLFormat, HiddenNames, Checking, ProjectError


'@section State
'===============================================================================

Private Type TCustomPivotTable
    sh As Worksheet
    shNames As HiddenNames
    hasChecks As Boolean
    sealed As Boolean
    checkCounter As Long
    internalChecks As Checking
End Type

Private this As TCustomPivotTable


'@section Constants
'===============================================================================

Private Const CLASS_NAME       As String = "CustomPivotTable"
Private Const PIVOT_PREFIX     As String = "PivotTable_"
Private Const HNME_OUTPUT_ROW  As String = "pivot_output_row"
Private Const HNME_COUNTER     As String = "pivot_counter"
Private Const HNME_HAS_CONTENT As String = "pivot_has_content"
Private Const HNME_FORMATTED   As String = "pivot_formatted"
Private Const TITLE_PREFIX     As String = "pivot_title_"
Private Const DEFAULT_START_ROW As Long = 2
Private Const TITLE_OFFSET     As Long = 2
Private Const PIVOT_OFFSET     As Long = 6
Private Const TITLE_COLUMN     As Long = 2

' Rows between one block's base and the next. A pivot is created empty and
' grows as the user drags fields into it, so the gap is what decides how big a
' pivot can get before it runs into the block below. Sizing it from the pivot's
' height at creation time would give every block three rows, which is why this
' is a fixed generous constant.
Private Const BLOCK_SIZE       As Long = 66

' RGB(155, 109, 88) - a Const cannot call RGB.
Private Const TITLE_FONT_COLOR As Long = 5795227


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new CustomPivotTable instance
'@details
'Validates the worksheet argument and initialises the worksheet-level
'HiddenNames listed under PUBLISHED STATE if they do not already exist. This
'allows multiple Create calls on the same sheet to resume from the current
'position.
'@param sh Worksheet. The dedicated pivot table worksheet in the output workbook.
'@return CustomPivotTable. The initialised instance.
'@throws InvalidArgument. When sh is Nothing.
'@export
Public Function Create(ByVal sh As Worksheet) As CustomPivotTable

    If sh Is Nothing Then ThrowError ProjectError.InvalidArgument, _
        "Unable to create custom pivot table, sheet is Nothing"

    Dim instance As CustomPivotTable
    Set instance = New CustomPivotTable
    Set instance.Wksh = sh
    instance.InitialiseMetadata
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes.
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:wksh-get
'@prop-title The dedicated pivot table worksheet
'@return Worksheet. The host worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the host worksheet (factory use only)
'@param sh Worksheet. The worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property


'@section Public API
'===============================================================================

'@label:add
'@sub-title Create a pivot table from a ListObject source and add a titled block
'@details
'Reads the current position from the worksheet-level HiddenNames, creates a
'PivotTable via PivotTableWizard, writes a styled title, stores the title text
'under pivot_title_<tableName>, and advances the position for the next block.
'
'The source table is resolved against every worksheet of the workbook before
'the wizard is called. A name that resolves to nothing is recorded as a
'checking and the wizard is skipped: PivotTableWizard answers such a name with
'a modal dialog, which On Error cannot see. That dialog freezes Excel in front
'of the user in the middle of a build and hangs a headless run. An error the
'wizard does raise is recorded the same way, so one bad table leaves the rest
'of the build alone. The position advances in both cases, so the next block
'cannot land on top of whatever the failed attempt left behind.
'@param title String. Display label for the pivot block (typically the sheet name).
'@param tableName String. Name of the source ListObject providing pivot data.
'@param pivotName String. Localised prefix for the title. Defaults to "Pivot Table".
'@throws ObjectNotInitialized. When the instance holds no worksheet.
'@export
Public Sub Add(ByVal title As String, _
               ByVal tableName As String, _
               Optional ByVal pivotName As String = "Pivot Table")

    Dim sh As Worksheet
    Dim shNames As HiddenNames
    Dim startRow As Long
    Dim counter As Long
    Dim titleCell As Range
    Dim outputCell As Range
    Dim titleText As String
    Dim pivotAdded As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    GuardReady "Add"

    Set sh = this.sh
    Set shNames = SheetNames()

    'Read current position
    startRow = shNames.ValueAsLong(HNME_OUTPUT_ROW, DEFAULT_START_ROW)
    counter = shNames.ValueAsLong(HNME_COUNTER, 1)

    'Compute ranges
    Set titleCell = sh.Cells(startRow + TITLE_OFFSET, TITLE_COLUMN)
    Set outputCell = sh.Cells(startRow + PIVOT_OFFSET, TITLE_COLUMN)

    'Create pivot table - a failure is logged, and the block is skipped. The
    'source is resolved first, because the wizard shows a dialog for a name it
    'cannot find and a dialog is out of reach of the handler below.
    If SourceTableExists(tableName) Then
        On Error GoTo PivotFailed
        sh.PivotTableWizard SourceType:=xlDatabase, _
                             SourceData:=tableName, _
                             tableName:=PIVOT_PREFIX & tableName, _
                             TableDestination:=outputCell
        On Error GoTo 0
        pivotAdded = True
    Else
        LogCheck tableName, "Unable to add pivot table " & tableName & _
                 ": the workbook holds no table of that name", checkingError
    End If

AdvanceBlock:
    On Error GoTo 0

    If errNumber <> 0 Then
        LogCheck tableName, "Unable to add pivot table " & tableName & ": " & _
                 errNumber & " - " & errDescription, checkingError
    End If

    If pivotAdded Then
        titleText = pivotName & " " & counter & " - " & title

        With titleCell
            .Value = titleText
            With .Font
                .Color = TITLE_FONT_COLOR
                .Italic = True
            End With
        End With

        'The linelist reads the title text back, so the name holds the text
        'itself.
        shNames.EnsureName TITLE_PREFIX & tableName, titleText, HiddenNameTypeString
        shNames.SetValue TITLE_PREFIX & tableName, titleText

        'The sheet has earned its visibility
        shNames.SetValue HNME_HAS_CONTENT, True
    End If

    'Advance position for next block, whether or not the pivot landed
    shNames.SetValue HNME_OUTPUT_ROW, startRow + BLOCK_SIZE
    shNames.SetValue HNME_COUNTER, counter + 1

    Exit Sub

PivotFailed:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume AdvanceBlock
End Sub

'@label:format
'@sub-title Apply design formatting to the pivot table worksheet
'@details
'Owns the sheet's visibility: a pivot sheet that never took a pivot stays
'very hidden, and one that did becomes visible. The cosmetic pass rewrites
'every row of the worksheet, so pivot_formatted holds it to one run however
'many times this is called.
'@param design LLFormat. The format object providing style rules.
'@throws ObjectNotInitialized. When the instance holds no worksheet.
'@throws InvalidArgument. When design is Nothing.
'@export
Public Sub Format(ByVal design As LLFormat)

    Dim shNames As HiddenNames

    GuardReady "Format"

    If design Is Nothing Then ThrowError ProjectError.InvalidArgument, _
        "Unable to format the custom pivot table sheet, design is Nothing"

    Set shNames = SheetNames()

    If Not shNames.ValueAsBoolean(HNME_HAS_CONTENT, False) Then
        this.sh.Visible = xlSheetVeryHidden
        Exit Sub
    End If

    this.sh.Visible = xlSheetVisible

    If shNames.ValueAsBoolean(HNME_FORMATTED, False) Then Exit Sub

    design.ApplyFormat this.sh, HListCustomPivotTables
    shNames.SetValue HNME_FORMATTED, True
End Sub

'@label:title
'@sub-title The stored title of one table's pivot block
'@details
'Answers what Add wrote for that table, which is what the linelist shows when
'it lists variables under their pivot. Empty when that table has no block.
'@param tableName String. Name of the source ListObject.
'@return String. The title text, or an empty string.
'@throws ObjectNotInitialized. When the instance holds no worksheet.
'@export
Public Function Title(ByVal tableName As String) As String
    GuardReady "Title"
    Title = SheetNames().ValueAsString(TITLE_PREFIX & tableName)
End Function


'@section Internal Helpers
'===============================================================================

'@label:initialise-metadata
'@sub-title Seed the worksheet-level HiddenNames this class reads
'@details
'Creates the four private names if they do not already exist. Called by the
'factory, so every later write can be a bare SetValue and re-wrapping a
'populated sheet resumes from the position it left.
Public Sub InitialiseMetadata()
    Dim shNames As HiddenNames

    Set shNames = SheetNames()

    If Not shNames.HasName(HNME_OUTPUT_ROW) Then _
        shNames.EnsureName HNME_OUTPUT_ROW, DEFAULT_START_ROW, HiddenNameTypeLong

    If Not shNames.HasName(HNME_COUNTER) Then _
        shNames.EnsureName HNME_COUNTER, 1, HiddenNameTypeLong

    If Not shNames.HasName(HNME_HAS_CONTENT) Then _
        shNames.EnsureName HNME_HAS_CONTENT, False, HiddenNameTypeBoolean

    If Not shNames.HasName(HNME_FORMATTED) Then _
        shNames.EnsureName HNME_FORMATTED, False, HiddenNameTypeBoolean
End Sub

'@label:sheet-names
'@sub-title The one HiddenNames manager for the host worksheet
'@details
'Creating a HiddenNames scans and caches every name in the container, so the
'instance is built once and held. Add, Format and InitialiseMetadata all read
'through this.
'@return HiddenNames. The cached manager.
Private Function SheetNames() As HiddenNames
    If this.shNames Is Nothing Then Set this.shNames = HiddenNames.Create(this.sh)
    Set SheetNames = this.shNames
End Function

'@label:source-table-exists
'@sub-title Whether the workbook defines the source ListObject
'@details
'PivotTableWizard shows a modal dialog for a SourceData naming no ListObject.
'In production that is a frozen Excel in the middle of a generation, and in a
'headless run it hangs the harness and leaves a zero-byte results file under a
'healthy import log. On Error reaches none of it, so the name is resolved here
'before the wizard is called.
'
'The source table of a pivot sits on the data entry sheet it was built from, so
'the walk covers every worksheet of the pivot sheet's workbook. It runs once
'per block added.
'@param tableName String. Name of the source ListObject.
'@return Boolean. True when one worksheet of the workbook defines that table.
Private Function SourceTableExists(ByVal tableName As String) As Boolean
    Dim candidateSheet As Worksheet
    Dim candidateTable As ListObject

    If tableName = vbNullString Then Exit Function

    For Each candidateSheet In this.sh.Parent.Worksheets
        For Each candidateTable In candidateSheet.ListObjects
            If StrComp(candidateTable.Name, tableName, vbTextCompare) = 0 Then
                SourceTableExists = True
                Exit Function
            End If
        Next candidateTable
    Next candidateSheet
End Function

'@label:guard-ready
'@sub-title Guard a member against use on an instance with no worksheet
'@param memberName String. The member name for the error message.
Private Sub GuardReady(ByVal memberName As String)
    If this.sh Is Nothing Then _
        ThrowError ProjectError.ObjectNotInitialized, _
                   CLASS_NAME & "." & memberName & " requires a worksheet. Use Create first."
End Sub


'@section Checkings
'===============================================================================

'@label:log-check
'@sub-title Append a diagnostic entry to the internal checking log
'@details
'The key carries the source table, so a reader of the generation report can see
'which pivot block the entry is about. The counter keeps the key free when the
'same table files twice, which `Checking.Add` requires.
'@param tableName String. Name of the source ListObject the entry is about.
'@param label String. Message describing the issue.
'@param scope Byte. Severity scope for the log entry.
Private Sub LogCheck(ByVal tableName As String, _
                     ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)

    If this.internalChecks Is Nothing Then _
        Set this.internalChecks = Checking.Create("Custom pivot tables")

    this.internalChecks.Add _
        CLASS_NAME & " " & tableName & " #" & CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether diagnostic entries have been recorded
'@return Boolean. True when at least one entry exists.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Checking. The checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error with class context
'@details
'Wraps Err.Raise with CLASS_NAME as the source. The ProjectError value is
'raised as it stands, the way every other class in the tree raises it.
'@param errNumber Long. The ProjectError constant to raise.
'@param errMessage String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errMessage As String)
    Err.Raise errNumber, CLASS_NAME, "Error: " & errMessage
End Sub
