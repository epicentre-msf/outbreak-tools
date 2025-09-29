VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionLayoutVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Section layout strategy for vertical linelist sheets")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSectionLayout

Private Type TVerticalLayout
    context As ILLSectionContext
    startRow As Long
    endRow As Long
    headerRange As Range
End Type

Private this As TVerticalLayout
Private Const CLASSNAME As String = "LLSectionLayoutVertical"
Private Const HEADER_COLUMN As Long = 2

'@section Guards
'===============================================================================
Private Sub EnsureConfigured()
    If this.context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Layout has not been configured with a section context"
    End If
    If this.startRow <= 0 Or this.endRow < this.startRow Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section layout start and end rows are not initialised"
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
                   "Unable to resolve variables for the requested section"
    End If
    Set VariableNames = names
End Function

Private Function ResolveRowIndex(ByVal varName As String) As Long
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
                   "Row index missing for variable '" & varName & "'"
    End If

    ResolveRowIndex = result
End Function

Private Function SectionWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set SectionWorksheet = this.context.Linelist.Wksh(sheetName)
    On Error GoTo 0
End Function

Private Sub BuildHeaderRange()
    Dim names As BetterArray
    Dim firstVar As String
    Dim lastVar As String
    Dim startRowSheet As Long
    Dim endRowSheet As Long
    Dim sheetName As String
    Dim sheet As Worksheet

    Set names = VariableNames()
    firstVar = CStr(names.Item(names.LowerBound))
    lastVar = CStr(names.Item(names.UpperBound))

    startRowSheet = ResolveRowIndex(firstVar)
    endRowSheet = ResolveRowIndex(lastVar)
    sheetName = this.context.Value("sheet name")

    Set sheet = SectionWorksheet(sheetName)
    If sheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Section worksheet '" & sheetName & "' was not found"
    End If

    Set this.headerRange = sheet.Range(sheet.Cells(startRowSheet, HEADER_COLUMN), _
                                       sheet.Cells(endRowSheet, HEADER_COLUMN))
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLSectionLayout_Configure(ByVal context As ILLSectionContext, _
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

    BuildHeaderRange
End Sub

Private Sub ILLSectionLayout_Reset()
    Set this.context = Nothing
    this.startRow = 0
    this.endRow = 0
    Set this.headerRange = Nothing
End Sub

Private Property Get ILLSectionLayout_HeaderRange() As Range
    EnsureConfigured
    Set ILLSectionLayout_HeaderRange = this.headerRange
End Property

Private Property Get ILLSectionLayout_PrintHeaderRange() As Range
    EnsureConfigured
    Set ILLSectionLayout_PrintHeaderRange = Nothing
End Property

Private Property Get ILLSectionLayout_CrfHeaderRange() As Range
    EnsureConfigured
    Set ILLSectionLayout_CrfHeaderRange = Nothing
End Property

Private Property Get ILLSectionLayout_LayoutKey() As String
    ILLSectionLayout_LayoutKey = "section-vertical"
End Property

