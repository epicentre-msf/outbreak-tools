VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableIterationItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable DTO capturing the outcome of a TableSpecsNavigator iteration")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableIterationItem

Private Type TIterationItemState
    specification As ITablesSpecs
    headerRange As Range
    rowRange As Range
    listRow As ListRow
    isNewSection As Boolean
    previousSpec As ITablesSpecs
    nextSpec As ITablesSpecs
End Type

Private this As TIterationItemState

Private Const CLASSNAME As String = "AnalysisTableIterationItem"

'@section Factory
'===============================================================================
'Create a fully initialised iteration item encapsulating a specification row.
Public Function Create(ByVal specification As ITablesSpecs, _
                       ByVal headerRange As Range, _
                       ByVal rowRange As Range, _
                       ByVal listRow As ListRow, _
                       ByVal isNewSection As Boolean, _
                       Optional ByVal previousSpec As ITablesSpecs = Nothing, _
                       Optional ByVal nextSpec As ITablesSpecs = Nothing) As IAnalysisTableIterationItem

    ValidateSpecification specification
    ValidateRowRange rowRange
    ValidateListRow listRow

    With New AnalysisTableIterationItem
        Set .Specification = specification
        Set .HeaderRange = headerRange
        Set .RowRange = rowRange
        Set .ListRow = listRow
        .IsNewSection = isNewSection
        Set .PreviousSpecification = previousSpec
        Set .NextSpecification = nextSpec
        Set Create = .Self
    End With
End Function

'Expose the instance through its interface to encourage loose coupling.
Public Property Get Self() As IAnalysisTableIterationItem
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'Assign the specification reference once validated.
Public Property Set Specification(ByVal specification As ITablesSpecs)
    ValidateSpecification specification
    Set this.specification = specification
End Property

'Return the wrapped specification reference.
Private Property Get SpecificationInternal() As ITablesSpecs
    Set SpecificationInternal = this.specification
End Property

'Store the header range used when building the specification.
Public Property Set HeaderRange(ByVal headerRange As Range)
    ValidateHeaderRange headerRange
    Set this.headerRange = headerRange
End Property

'Return the cached header range reference.
Private Property Get HeaderRangeInternal() As Range
    Set HeaderRangeInternal = this.headerRange
End Property

'Store the row range used to create the specification.
Public Property Set RowRange(ByVal rowRange As Range)
    ValidateRowRange rowRange
    Set this.rowRange = rowRange
End Property

'Access the stored row range.
Private Property Get RowRangeInternal() As Range
    Set RowRangeInternal = this.rowRange
End Property

'Store the originating list row.
Public Property Set ListRow(ByVal listRow As ListRow)
    ValidateListRow listRow
    Set this.listRow = listRow
End Property

'Access the originating list row instance.
Private Property Get ListRowInternal() As ListRow
    Set ListRowInternal = this.listRow
End Property

'Record whether the current specification starts a new section.
Public Property Let IsNewSection(ByVal isNewSection As Boolean)
    this.isNewSection = isNewSection
End Property

'Expose the recorded section flag.
Private Property Get IsNewSectionInternal() As Boolean
    IsNewSectionInternal = this.isNewSection
End Property

'Store the previous specification reference when available.
Public Property Set PreviousSpecification(ByVal previousSpec As ITablesSpecs)
    Set this.previousSpec = previousSpec
End Property

'Return the recorded previous specification reference.
Private Property Get PreviousSpecificationInternal() As ITablesSpecs
    Set PreviousSpecificationInternal = this.previousSpec
End Property

'Store the next specification reference when available.
Public Property Set NextSpecification(ByVal nextSpec As ITablesSpecs)
    Set this.nextSpec = nextSpec
End Property

'Access the recorded next specification reference.
Private Property Get NextSpecificationInternal() As ITablesSpecs
    Set NextSpecificationInternal = this.nextSpec
End Property

'@section Validation helpers
'===============================================================================
'Ensure the provided specification reference is usable.
Private Sub ValidateSpecification(ByVal specification As ITablesSpecs)
    If specification Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "specification cannot be Nothing"
    End If

    If LenB(Trim$(TypeName(specification))) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), _
                   "specification must expose a concrete implementation"
    End If
End Sub

'Ensure the supplied header range is valid.
Private Sub ValidateHeaderRange(ByVal headerRange As Range)
    If headerRange Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "headerRange cannot be Nothing"
    End If

    If StrComp(TypeName(headerRange), "Range", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "headerRange must be a Range instance"
    End If
End Sub

'Ensure the supplied range corresponds to a valid table row.
Private Sub ValidateRowRange(ByVal rowRange As Range)
    If rowRange Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "rowRange cannot be Nothing"
    End If

    If StrComp(TypeName(rowRange), "Range", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "rowRange must be a Range instance"
    End If
End Sub

'Ensure the provided list row reference is valid.
Private Sub ValidateListRow(ByVal listRow As ListRow)
    If listRow Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "listRow cannot be Nothing"
    End If

    If StrComp(TypeName(listRow), "ListRow", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "listRow must be a ListRow instance"
    End If
End Sub

'Centralised helper for raising consistent project errors.
Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisTableIterationItem_Specification() As ITablesSpecs
    Set IAnalysisTableIterationItem_Specification = SpecificationInternal()
End Property

Private Property Get IAnalysisTableIterationItem_IsNewSection() As Boolean
    IAnalysisTableIterationItem_IsNewSection = IsNewSectionInternal()
End Property

Private Property Get IAnalysisTableIterationItem_HeaderRange() As Range
    Set IAnalysisTableIterationItem_HeaderRange = HeaderRangeInternal()
End Property

Private Property Get IAnalysisTableIterationItem_RowRange() As Range
    Set IAnalysisTableIterationItem_RowRange = RowRangeInternal()
End Property

Private Property Get IAnalysisTableIterationItem_ListRow() As ListRow
    Set IAnalysisTableIterationItem_ListRow = ListRowInternal()
End Property

Private Property Get IAnalysisTableIterationItem_PreviousSpecification() As ITablesSpecs
    Set IAnalysisTableIterationItem_PreviousSpecification = PreviousSpecificationInternal()
End Property

Private Property Get IAnalysisTableIterationItem_NextSpecification() As ITablesSpecs
    Set IAnalysisTableIterationItem_NextSpecification = NextSpecificationInternal()
End Property
