VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSubSectionLayoutHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Subsection layout strategy for horizontal linelist sheets")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSubSectionLayout

Private Type THorizontalSubLayout
    context As ILLSectionContext
    startRow As Long
    endRow As Long
    headerRange As Range
    printHeaderRange As Range
    crfHeaderRange As Range
End Type

Private this As THorizontalSubLayout
Private Const CLASSNAME As String = "LLSubSectionLayoutHorizontal"
Private Const HEADER_ROW As Long = 6
Private Const CRF_COLUMN_END As Long = 7

'@section Guards
'===============================================================================
Private Sub EnsureConfigured()
    If this.context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Layout has not been configured with a section context"
    End If
    If this.startRow <= 0 Or this.endRow < this.startRow Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Subsection layout start and end rows are not initialised"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Helpers
'===============================================================================
Private Function VariableNames() As BetterArray
    Dim names As BetterArray
    Set names = this.context.VariableNames(this.startRow, this.endRow)
    If names Is Nothing Or names.Length = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Unable to resolve variables for the requested subsection"
    End If
    Set VariableNames = names
End Function

Private Function ResolveColumnIndex(ByVal varName As String) As Long
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim valueText As String
    Dim result As Long

    Set dict = this.context.Dictionary
    Set vars = LLVariables.Create(dict)

    valueText = vars.Value(colName:="column index", varName:=varName)
    On Error Resume Next
        result = CLng(valueText)
    On Error GoTo 0

    If result = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column index missing for variable '" & varName & "'"
    End If

    ResolveColumnIndex = result
End Function

Private Function ResolveCrfHeaderRow(ByVal varName As String) As Long
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim valueText As String
    Dim crfLine As Long

    Set dict = this.context.Dictionary
    Set vars = LLVariables.Create(dict)

    valueText = vars.Value(colName:="crf index", varName:=varName)
    On Error Resume Next
        crfLine = CLng(valueText)
    On Error GoTo 0

    If crfLine <= 3 Then
        ResolveCrfHeaderRow = 0
    Else
        ResolveCrfHeaderRow = crfLine - 3
    End If
End Function

Private Function SectionWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set SectionWorksheet = this.context.Linelist.Wksh(sheetName)
    On Error GoTo 0
End Function

Private Function PrintWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set PrintWorksheet = this.context.Linelist.PrintWksh(sheetName)
    On Error GoTo 0
End Function

Private Function CrfWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set CrfWorksheet = this.context.Linelist.CRFWksh(sheetName)
    On Error GoTo 0
End Function

Private Sub BuildHeaderRanges()
    Dim names As BetterArray
    Dim firstVar As String
    Dim lastVar As String
    Dim startColumn As Long
    Dim endColumn As Long
    Dim sheetName As String
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim crfSheet As Worksheet
    Dim crfHeaderRow As Long

    Set names = VariableNames()
    firstVar = CStr(names.Item(names.LowerBound))
    lastVar = CStr(names.Item(names.UpperBound))

    startColumn = ResolveColumnIndex(firstVar)
    endColumn = ResolveColumnIndex(lastVar)
    sheetName = this.context.Value("sheet name")

    Set sheet = SectionWorksheet(sheetName)
    If sheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Subsection worksheet '" & sheetName & "' was not found"
    End If

    Set this.headerRange = sheet.Range(sheet.Cells(HEADER_ROW, startColumn), _
                                       sheet.Cells(HEADER_ROW, endColumn))

    Set printSheet = PrintWorksheet(sheetName)
    If Not printSheet Is Nothing Then
        Set this.printHeaderRange = printSheet.Range(printSheet.Cells(HEADER_ROW, startColumn), _
                                                     printSheet.Cells(HEADER_ROW, endColumn))
    Else
        Set this.printHeaderRange = Nothing
    End If

    crfHeaderRow = ResolveCrfHeaderRow(firstVar)
    If crfHeaderRow > 0 Then
        Set crfSheet = CrfWorksheet(sheetName)
        If Not crfSheet Is Nothing Then
            Set this.crfHeaderRange = crfSheet.Range(crfSheet.Cells(crfHeaderRow, 1), _
                                                     crfSheet.Cells(crfHeaderRow, CRF_COLUMN_END))
        Else
            Set this.crfHeaderRange = Nothing
        End If
    Else
        Set this.crfHeaderRange = Nothing
    End If
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLSubSectionLayout_Configure(ByVal context As ILLSectionContext, _
                                          ByVal startRow As Long, _
                                          ByVal endRow As Long)

    If context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Section context reference is required"
    End If

    If startRow <= 0 Or endRow < startRow Then
        ThrowError ProjectError.InvalidArgument, _
                   "Start row and end row must describe a positive range"
    End If

    Set this.context = context
    this.startRow = startRow
    this.endRow = endRow
    Set this.headerRange = Nothing
    Set this.printHeaderRange = Nothing
    Set this.crfHeaderRange = Nothing

    BuildHeaderRanges
End Sub

Private Sub ILLSubSectionLayout_Reset()
    Set this.context = Nothing
    this.startRow = 0
    this.endRow = 0
    Set this.headerRange = Nothing
    Set this.printHeaderRange = Nothing
    Set this.crfHeaderRange = Nothing
End Sub

Private Property Get ILLSubSectionLayout_HeaderRange() As Range
    EnsureConfigured
    Set ILLSubSectionLayout_HeaderRange = this.headerRange
End Property

Private Property Get ILLSubSectionLayout_PrintHeaderRange() As Range
    EnsureConfigured
    If this.printHeaderRange Is Nothing Then Exit Property
    Set ILLSubSectionLayout_PrintHeaderRange = this.printHeaderRange
End Property

Private Property Get ILLSubSectionLayout_CrfHeaderRange() As Range
    EnsureConfigured
    If this.crfHeaderRange Is Nothing Then Exit Property
    Set ILLSubSectionLayout_CrfHeaderRange = this.crfHeaderRange
End Property

Private Property Get ILLSubSectionLayout_LayoutKey() As String
    ILLSubSectionLayout_LayoutKey = "subsection-horizontal"
End Property

