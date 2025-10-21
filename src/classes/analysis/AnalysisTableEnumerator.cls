VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Enumerates analysis table specifications using TableSpecsNavigator to provide rich traversal context")
'@depends ProjectError, BetterArray, TableSpecsNavigator, TableSpecsColumnMap, AnalysisTableSpecBuilder, AnalysisTableIterationItem
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableEnumerator

Private Type TEnumeratorState
    builder As ITableSpecsBuilder
End Type

Private this As TEnumeratorState

Private Const CLASSNAME As String = "AnalysisTableEnumerator"

'@section Factory
'===============================================================================
'Create an enumerator optionally configured with a custom spec builder.
Public Function Create(Optional ByVal builder As ITableSpecsBuilder = Nothing) As IAnalysisTableEnumerator
    With New AnalysisTableEnumerator
        If builder Is Nothing Then
            Set .Builder = AnalysisTableSpecBuilder.Create
        Else
            Set .Builder = builder
        End If
        Set Create = .Self
    End With
End Function

'Expose the current instance through its interface.
Public Property Get Self() As IAnalysisTableEnumerator
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'Assign the spec builder dependency after validation.
Public Property Set Builder(ByVal builder As ITableSpecsBuilder)
    ValidateBuilder builder
    Set this.builder = builder
End Property

'Access the configured builder dependency.
Private Property Get BuilderInternal() As ITableSpecsBuilder
    Set BuilderInternal = this.builder
End Property

'@section Enumeration helpers
'===============================================================================
'Materialise the anchor range used to limit navigator lookups.
Private Function ResolveAnchorRange(ByVal specificationList As ListObject) As Range
    Dim lastRow As ListRow

    If specificationList.ListRows.Count = 0 Then Exit Function

    Set lastRow = specificationList.ListRows(specificationList.ListRows.Count)
    Set ResolveAnchorRange = lastRow.Range
End Function

'Build a column map tailored for the provided list row.
Private Function CreateColumnMap(ByVal headerRange As Range, ByVal rowRange As Range) As ITableSpecsColumnMap
    Set CreateColumnMap = TableSpecsColumnMap.Create(headerRange, rowRange)
End Function

'Instantiate a navigator positioned on the supplied list row.
Private Function CreateNavigator(ByVal headerRange As Range, _
                                 ByVal rowRange As Range, _
                                 ByVal linelistSpecs As ILinelistSpecs, _
                                 ByVal columnMap As ITableSpecsColumnMap, _
                                 ByVal tableTypeCode As Byte) As ITableSpecsNavigator

    Set CreateNavigator = TableSpecsNavigator.Create(headerRange, rowRange, linelistSpecs, columnMap, BuilderInternal(), tableTypeCode)
End Function

'Safely locate the previous specification if the navigator is not positioned on a new section.
Private Function ResolvePreviousSpec(ByVal navigator As ITableSpecsNavigator, ByVal isNewSection As Boolean) As ITablesSpecs
    Dim candidate As ITablesSpecs

    Set candidate = Nothing

    If isNewSection Then Exit Function

    On Error Resume Next
        Set candidate = navigator.PreviousSpec
        If Err.Number <> 0 Then
            Set candidate = Nothing
            Err.Clear
        End If
    On Error GoTo 0

    Set ResolvePreviousSpec = candidate
End Function

'Safely locate the next specification bounded by the provided anchor.
Private Function ResolveNextSpec(ByVal navigator As ITableSpecsNavigator, ByVal anchorRange As Range) As ITablesSpecs
    Dim candidate As ITablesSpecs

    Set candidate = Nothing

    If anchorRange Is Nothing Then Exit Function

    On Error Resume Next
        Set candidate = navigator.NextSpec(anchorRange)
        If Err.Number <> 0 Then
            Set candidate = Nothing
            Err.Clear
        End If
    On Error GoTo 0

    Set ResolveNextSpec = candidate
End Function

'Construct a result item capturing the traversal context for the current row.
Private Function CreateIterationItem(ByVal specification As ITablesSpecs, _
                                     ByVal headerRange As Range, _
                                     ByVal rowRange As Range, _
                                     ByVal listRow As ListRow, _
                                     ByVal isNewSection As Boolean, _
                                     ByVal previousSpec As ITablesSpecs, _
                                     ByVal nextSpec As ITablesSpecs) As IAnalysisTableIterationItem
    Set CreateIterationItem = AnalysisTableIterationItem.Create(specification, headerRange, rowRange, listRow, isNewSection, previousSpec, nextSpec)
End Function

'Create an empty BetterArray configured with a base lower bound for predictable results.
Private Function CreateResultBuffer() As BetterArray
    Dim buffer As BetterArray

    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateResultBuffer = buffer
End Function

'@section Validation helpers
'===============================================================================
'Ensure the listobject parameter is valid.
Private Sub ValidateSpecificationList(ByVal specificationList As ListObject)
    If specificationList Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "specificationList cannot be Nothing"
    End If

    If StrComp(TypeName(specificationList), "ListObject", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), _
                   "specificationList must be a ListObject instance"
    End If
End Sub

'Ensure the linelist reference is valid.
Private Sub ValidateLinelist(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelistSpecs cannot be Nothing"
    End If

    If LenB(Trim$(TypeName(linelistSpecs))) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), _
                   "linelistSpecs must expose a concrete implementation"
    End If
End Sub

'Ensure the table type code is within a reasonable range.
Private Sub ValidateTableTypeCode(ByVal tableTypeCode As Byte)
    If tableTypeCode = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "tableTypeCode must be greater than zero"
    End If
End Sub

'Ensure a usable builder dependency has been configured.
Private Sub ValidateBuilder(ByVal builder As ITableSpecsBuilder)
    If builder Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "builder cannot be Nothing"
    End If

    If LenB(Trim$(TypeName(builder))) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), _
                   "builder must expose a concrete implementation"
    End If
End Sub

'Centralised helper for raising consistent project errors.
Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTableEnumerator_Enumerate(ByVal specificationList As ListObject, _
                                                    ByVal linelistSpecs As ILinelistSpecs, _
                                                    ByVal tableTypeCode As Byte) As BetterArray
    Dim results As BetterArray
    Dim headerRange As Range
    Dim anchorRange As Range
    Dim listRow As ListRow
    Dim rowRange As Range
    Dim columnMap As ITableSpecsColumnMap
    Dim navigator As ITableSpecsNavigator
    Dim specification As ITablesSpecs
    Dim isNewSection As Boolean
    Dim previousSpec As ITablesSpecs
    Dim nextSpec As ITablesSpecs
    Dim builder As ITableSpecsBuilder

    ValidateSpecificationList specificationList
    ValidateLinelist linelistSpecs
    ValidateTableTypeCode tableTypeCode

    Set results = CreateResultBuffer()

    If specificationList.DataBodyRange Is Nothing Then
        Set IAnalysisTableEnumerator_Enumerate = results
        Exit Function
    End If

    Set headerRange = specificationList.HeaderRowRange
    Set anchorRange = ResolveAnchorRange(specificationList)
    Set builder = BuilderInternal

    For Each listRow In specificationList.ListRows
        Set rowRange = listRow.Range
        Set columnMap = CreateColumnMap(headerRange, rowRange)
        Set navigator = CreateNavigator(headerRange, rowRange, linelistSpecs, columnMap, tableTypeCode)
        Set specification = builder.CreateSpec(headerRange, rowRange, linelistSpecs)

        If specification Is Nothing Then GoTo ContinueLoop
        If Not specification.ValidTable Then GoTo ContinueLoop

        isNewSection = navigator.IsNewSection
        Set previousSpec = ResolvePreviousSpec(navigator, isNewSection)
        Set nextSpec = ResolveNextSpec(navigator, anchorRange)

        results.Push CreateIterationItem(specification, headerRange, rowRange, listRow, isNewSection, previousSpec, nextSpec)

ContinueLoop:
        Set navigator = Nothing
        Set columnMap = Nothing
        Set specification = Nothing
    Next listRow

    Set IAnalysisTableEnumerator_Enumerate = results
End Function
