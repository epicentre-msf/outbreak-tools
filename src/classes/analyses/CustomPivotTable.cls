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

Implements ICustomPivotTable

'@class CustomPivotTable
'@description
'Manages pivot table creation on a dedicated worksheet during linelist
'building. Each call to Add creates a new PivotTable block (title + pivot)
'stacked vertically. Position tracking and title references are stored as
'worksheet-level HiddenNames instead of hidden cell values.
'
'@depends ICustomPivotTable, ILLFormat, IHiddenNames, HiddenNames, ProjectError


'@section State
'===============================================================================

Private Type TCustomPivotTable
    sh As Worksheet
End Type

Private this As TCustomPivotTable


'@section Constants
'===============================================================================

Private Const CLASS_NAME       As String = "CustomPivotTable"
Private Const PIVOT_PREFIX     As String = "PivotTable_"
Private Const HNME_OUTPUT_ROW  As String = "pivot_output_row"
Private Const HNME_COUNTER     As String = "pivot_counter"
Private Const TITLE_PREFIX     As String = "RNG_PivotTitle_"
Private Const DEFAULT_START_ROW As Long = 2
Private Const TITLE_OFFSET     As Long = 2
Private Const PIVOT_OFFSET     As Long = 6
Private Const BLOCK_SIZE       As Long = 60


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new CustomPivotTable instance
'@details
'Validates the worksheet argument and initialises worksheet-level HiddenNames
'for position tracking (pivot_output_row, pivot_counter) if they do not
'already exist. This allows multiple Create calls on the same sheet to
'resume from the current position.
'@param sh Worksheet. The dedicated pivot table worksheet in the output workbook.
'@return ICustomPivotTable. The initialised instance.
'@throws InvalidArgument. When sh is Nothing.
'@export
Public Function Create(ByVal sh As Worksheet) As ICustomPivotTable

    If sh Is Nothing Then ThrowError InvalidArgument, _
        "Unable to create custom pivot table, sheet is Nothing"

    With New CustomPivotTable
        Set .Wksh = sh
        .InitialiseMetadata
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Return the current instance as ICustomPivotTable
'@return ICustomPivotTable. Self-reference for factory pattern.
Public Property Get Self() As ICustomPivotTable
    Set Self = Me
End Property

'@label:wksh-get
'@prop-title The dedicated pivot table worksheet
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the host worksheet (factory use only)
'@param sh Worksheet. The worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property


'@section Public API
'===============================================================================

'@label:add
'@sub-title Create a pivot table from a ListObject source and add a titled block
'@details
'Reads the current position from worksheet-level HiddenNames, creates a
'PivotTable via PivotTableWizard, writes a styled title, stores the title
'cell address as a HiddenName, and advances the position for the next block.
'If the PivotTableWizard call fails (e.g. invalid source), the error is
'logged to Debug.Print and the method exits without creating a title.
'@param title String. Display label for the pivot block (typically the sheet name).
'@param tableName String. Name of the source ListObject providing pivot data.
'@param pivotName String. Localised prefix for the title. Defaults to "Pivot Table".
Public Sub Add(ByVal title As String, _
               ByVal tableName As String, _
               Optional ByVal pivotName As String = "Pivot Table")

    Dim sh As Worksheet
    Dim shNames As IHiddenNames
    Dim startRow As Long
    Dim counter As Long
    Dim titleCell As Range
    Dim outputCell As Range

    Set sh = this.sh
    Set shNames = HiddenNames.Create(sh)

    'Read current position
    startRow = CLng(shNames.ValueAsString(HNME_OUTPUT_ROW))
    counter = CLng(shNames.ValueAsString(HNME_COUNTER))

    'Compute ranges
    Set titleCell = sh.Cells(startRow + TITLE_OFFSET, 2)
    Set outputCell = sh.Cells(startRow + PIVOT_OFFSET, 2)

    'Create pivot table — exit gracefully on failure
    On Error GoTo ErrAdding
    sh.PivotTableWizard SourceType:=xlDatabase, _
                         SourceData:=tableName, _
                         tableName:=PIVOT_PREFIX & tableName, _
                         TableDestination:=outputCell
    On Error GoTo 0

    'Write title
    titleCell.Value = pivotName & " " & counter & " - " & title
    titleCell.Font.Color = RGB(155, 109, 88)
    titleCell.Font.Italic = True

    'Store title cell address as worksheet-level HiddenName
    shNames.EnsureName TITLE_PREFIX & tableName, titleCell.Address, HiddenNameTypeString

    'Advance position for next block
    shNames.EnsureName HNME_OUTPUT_ROW, CStr(startRow + BLOCK_SIZE), HiddenNameTypeLong
    shNames.SetValue HNME_OUTPUT_ROW, CStr(startRow + BLOCK_SIZE)
    shNames.EnsureName HNME_COUNTER, CStr(counter + 1), HiddenNameTypeLong
    shNames.SetValue HNME_COUNTER, CStr(counter + 1)

    Exit Sub

ErrAdding:
    Debug.Print "Unable to add pivot table " & tableName & ": " & _
                Err.Number & " - " & Err.Description
End Sub

'@label:format
'@sub-title Apply design formatting to the pivot table worksheet
'@details
'Delegates to ILLFormat.ApplyFormat with the HListCustomPivotTables scope.
'@param design ILLFormat. The format object providing style rules.
Public Sub Format(ByVal design As ILLFormat)
    design.ApplyFormat this.sh, HListCustomPivotTables
End Sub


'@section Internal Helpers
'===============================================================================

'@label:initialise-metadata
'@sub-title Initialise worksheet-level HiddenNames for position tracking
'@details
'Creates pivot_output_row and pivot_counter HiddenNames if they do not
'already exist. Called by the factory to allow resuming from the current
'position when the same sheet is re-wrapped.
Public Sub InitialiseMetadata()
    Dim shNames As IHiddenNames
    Set shNames = HiddenNames.Create(this.sh)

    If Not shNames.HasName(HNME_OUTPUT_ROW) Then
        shNames.EnsureName HNME_OUTPUT_ROW, CStr(DEFAULT_START_ROW), HiddenNameTypeLong
    End If

    If Not shNames.HasName(HNME_COUNTER) Then
        shNames.EnsureName HNME_COUNTER, CStr(1), HiddenNameTypeLong
    End If
End Sub


'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error with class context
'@details
'Wraps Err.Raise with CLASS_NAME as the source, providing consistent
'error attribution across all methods in this class.
'@param errNumber Long. The ProjectError constant to raise.
'@param errMessage String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errMessage As String)
    Err.Raise VBA.vbObjectError + errNumber, CLASS_NAME, "Error: " & errMessage
End Sub


'@section Interface Implementation
'===============================================================================

Private Sub ICustomPivotTable_Add(ByVal title As String, _
                                   ByVal tableName As String, _
                                   Optional ByVal pivotName As String = "Pivot Table")
    Add title, tableName, pivotName
End Sub

Private Sub ICustomPivotTable_Format(ByVal design As ILLFormat)
    Format design
End Sub
