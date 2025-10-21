VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Iterator responsible for traversing TableSpecs rows")
'@depends ProjectError, TableSpecsPolicyHelpers, TableSpecsColumnMap
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsNavigator

Private Type TNavigatorState
    headerRange As Range
    currentRange As Range
    linelistSpecs As ILinelistSpecs
    columnMap As ITableSpecsColumnMap
    builder As ITableSpecsBuilder
    tableTypeCode As Byte
End Type

Private this As TNavigatorState
Private Const CLASSNAME As String = "TableSpecsNavigator"

'@section Factory
'===============================================================================
'Initialise the navigator with the provided context.
Public Function Create(ByVal headerRange As Range, _
                       ByVal currentRange As Range, _
                       ByVal linelistSpecs As ILinelistSpecs, _
                       ByVal columnMap As ITableSpecsColumnMap, _
                       ByVal builder As ITableSpecsBuilder, _
                       ByVal tableTypeCode As Byte) As ITableSpecsNavigator

    ValidateRange headerRange, "headerRange"
    ValidateRange currentRange, "currentRange"
    ValidateLinelistSpecs linelistSpecs
    ValidateColumnMap columnMap
    ValidateBuilder builder

    With New TableSpecsNavigator
        Set .HeaderRange = headerRange
        Set .CurrentRange = currentRange
        Set .LinelistSpecifications = linelistSpecs
        Set .ColumnMap = columnMap
        Set .Builder = builder
        .TableTypeCode = tableTypeCode
        columnMap.Refresh
        Set Create = .Self
    End With
End Function

'Expose the current instance as its interface type.
Public Property Get Self() As ITableSpecsNavigator
    Set Self = Me
End Property

'@section Context configuration
'===============================================================================
Public Property Set HeaderRange(ByVal headerRange As Range)
    ValidateRange headerRange, "headerRange"
    Set this.headerRange = headerRange
End Property

Public Property Set CurrentRange(ByVal currentRange As Range)
    ValidateRange currentRange, "currentRange"
    Set this.currentRange = currentRange
End Property

Public Property Set LinelistSpecifications(ByVal linelistSpecs As ILinelistSpecs)
    ValidateLinelistSpecs linelistSpecs
    Set this.linelistSpecs = linelistSpecs
End Property

Public Property Set ColumnMap(ByVal columnMap As ITableSpecsColumnMap)
    ValidateColumnMap columnMap
    Set this.columnMap = columnMap
End Property

Public Property Set Builder(ByVal builder As ITableSpecsBuilder)
    ValidateBuilder builder
    Set this.builder = builder
End Property

Public Property Let TableTypeCode(ByVal tableTypeCode As Byte)
    this.tableTypeCode = tableTypeCode
End Property

'@section Navigation helpers
'===============================================================================
Public Function PreviousSpec() As ITablesSpecs
    Dim candidateRange As Range
    Dim candidateSpec As ITablesSpecs

    EnsureStateReady

    If IsNewSectionInternal(this.currentRange) Then
        ThrowError CLng(ProjectError.ElementNotFound), _
                   "Current row starts a new section; no previous specification is available"
    End If

    Set candidateRange = this.currentRange.Offset(-1, 0)

    Do While Not candidateRange Is Nothing
        If candidateRange.Row <= this.headerRange.Row Then Exit Do

        Set candidateSpec = CreateSpecForRange(candidateRange)
        If candidateSpec Is Nothing Then Exit Do
        If candidateSpec.ValidTable Then Exit Do
        If IsNewSectionInternal(candidateRange) Then Exit Do

        Set candidateRange = candidateRange.Offset(-1, 0)
    Loop

    Set PreviousSpec = candidateSpec
End Function

Public Function NextSpec(ByVal anchorRange As Range) As ITablesSpecs
    Dim candidateRange As Range
    Dim candidateSpec As ITablesSpecs

    ValidateRange anchorRange, "anchorRange"
    EnsureStateReady

    Set candidateRange = this.currentRange.Offset(1, 0)

    Do While Not candidateRange Is Nothing
        If anchorRange Is Nothing Then Exit Do
        If candidateRange.Row > anchorRange.Row Then Exit Do

        Set candidateSpec = CreateSpecForRange(candidateRange)
        If Not candidateSpec Is Nothing Then
            If candidateSpec.ValidTable Then
                Set NextSpec = candidateSpec
                Exit Function
            End If
        End If

        Set candidateRange = candidateRange.Offset(1, 0)
    Loop
End Function

Public Function IsNewSection() As Boolean
    EnsureStateReady
    IsNewSection = IsNewSectionInternal(this.currentRange)
End Function

Public Sub Invalidate()
    If Not this.columnMap Is Nothing Then this.columnMap.Invalidate
End Sub

'@section Internal helpers
'===============================================================================
Private Sub EnsureStateReady()
    ValidateRange this.headerRange, "headerRange"
    ValidateRange this.currentRange, "currentRange"
    ValidateLinelistSpecs this.linelistSpecs
    ValidateColumnMap this.columnMap
    ValidateBuilder this.builder
End Sub

Private Function CreateSpecForRange(ByVal tableRange As Range) As ITablesSpecs
    On Error GoTo CleanFail
        Set CreateSpecForRange = this.builder.CreateSpec(this.headerRange, tableRange, this.linelistSpecs)
    On Error GoTo 0
    If CreateSpecForRange Is Nothing Then
        ThrowError CLng(ProjectError.ErrorUnexpectedState), _
                   "Builder returned Nothing for row " & CStr(tableRange.Row)
    End If
    Exit Function
CleanFail:
    ThrowError CLng(ProjectError.ErrorUnexpectedState), _
               "Failed to create TableSpecs instance for row " & CStr(tableRange.Row)
End Function

Private Function IsNewSectionInternal(ByVal tableRange As Range) As Boolean
    Dim sectionExists As Boolean
    Dim columnIndex As Long
    Dim currentSection As String
    Dim previousSection As String
    Dim previousCell As Range
    Dim localMap As ITableSpecsColumnMap

    IsNewSectionInternal = False

    If this.tableTypeCode = TABLE_TYPE_GLOBAL_SUMMARY Then Exit Function

    If tableRange Is this.currentRange Then
        Set localMap = this.columnMap
    Else
        On Error GoTo CleanFail
            Set localMap = TableSpecsColumnMap.Create(this.headerRange, tableRange)
        On Error GoTo 0
    End If

    sectionExists = localMap.ColumnExists("section")
    If Not sectionExists Then
        IsNewSectionInternal = True
        Exit Function
    End If

    columnIndex = localMap.ColumnIndex("section")
    If columnIndex <= 0 Then
        IsNewSectionInternal = True
        Exit Function
    End If

    currentSection = NormalizeValue(localMap.Value("section"))

    On Error Resume Next
        Set previousCell = tableRange.Cells(1, columnIndex).Offset(-1, 0)
    On Error GoTo 0

    If previousCell Is Nothing Then
        IsNewSectionInternal = True
        Exit Function
    End If

    previousSection = NormalizeValue(CStr(previousCell.Value))

    IsNewSectionInternal = Not ((currentSection = previousSection) And _
                               (previousCell.Row <> this.headerRange.Row))
    Exit Function

CleanFail:
    ThrowError CLng(ProjectError.ErrorUnexpectedState), _
               "Unable to evaluate section information"
End Function

Private Sub ValidateRange(ByVal candidate As Range, ByVal argumentName As String)
    If candidate Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
    If StrComp(TypeName(candidate), "Range", vbBinaryCompare) <> 0 Then
        ThrowError CLng(ProjectError.InvalidArgument), argumentName & " must be a valid Range reference"
    End If
End Sub

Private Sub ValidateLinelistSpecs(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "linelistSpecs cannot be Nothing"
    End If
End Sub

Private Sub ValidateColumnMap(ByVal columnMap As ITableSpecsColumnMap)
    If columnMap Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "columnMap cannot be Nothing"
    End If
End Sub

Private Sub ValidateBuilder(ByVal builder As ITableSpecsBuilder)
    If builder Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "builder cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsNavigator_Create(ByVal headerRange As Range, _
                                             ByVal currentRange As Range, _
                                             ByVal linelistSpecs As ILinelistSpecs, _
                                             ByVal columnMap As ITableSpecsColumnMap, _
                                             ByVal builder As ITableSpecsBuilder, _
                                             ByVal tableTypeCode As Byte) As ITableSpecsNavigator
    Set ITableSpecsNavigator_Create = Create(headerRange, currentRange, linelistSpecs, columnMap, builder, tableTypeCode)
End Function

Private Function ITableSpecsNavigator_PreviousSpec() As ITablesSpecs
    Set ITableSpecsNavigator_PreviousSpec = PreviousSpec()
End Function

Private Function ITableSpecsNavigator_NextSpec(ByVal anchorRange As Range) As ITablesSpecs
    Set ITableSpecsNavigator_NextSpec = NextSpec(anchorRange)
End Function

Private Function ITableSpecsNavigator_IsNewSection() As Boolean
    ITableSpecsNavigator_IsNewSection = IsNewSection()
End Function

Private Sub ITableSpecsNavigator_Invalidate()
    Invalidate
End Sub

