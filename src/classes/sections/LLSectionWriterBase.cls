VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionWriterBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Reusable writer orchestrating linelist sections")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSectionWriter

Private Type TSectionWriter
    context As ILLSectionContext
    layout As ILLSectionLayout
    hooks As ILLSectionWriterHooks
    isReady As Boolean
End Type

Private this As TSectionWriter
Private Const CLASSNAME As String = "LLSectionWriterBase"

'@section Guard helpers
'===============================================================================
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Sub EnsureReady()
    If Not this.isReady Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section writer has not been initialised"
    End If
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLSectionWriter_Initialise(ByVal context As ILLSectionContext, _
                                        ByVal layout As ILLSectionLayout, _
                                        ByVal hooks As ILLSectionWriterHooks)

    If context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Section context reference is required"
    End If

    If layout Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Section layout reference is required"
    End If

    If hooks Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Section writer hooks are required"
    End If

    Set this.context = context
    Set this.layout = layout
    Set this.hooks = hooks
    this.isReady = True
End Sub

Private Sub ILLSectionWriter_Reset()
    Set this.context = Nothing
    Set this.layout = Nothing
    Set this.hooks = Nothing
    this.isReady = False
End Sub

Private Sub ILLSectionWriter_WriteSection(ByVal startRow As Long, ByVal endRow As Long)
    EnsureReady

    If startRow <= 0 Or endRow < startRow Then
        ThrowError ProjectError.InvalidArgument, _
                   "Invalid section boundaries supplied"
    End If

    Dim originalStartRow As Long
    Dim sectionName As String
    Dim design As ILLFormat
    Dim headerRange As Range
    Dim printRange As Range
    Dim crfRange As Range
    Dim boundaryColumns As Variant
    Dim currentRow As Long
    Dim subsectionEnd As Long
    Dim subLayout As ILLSubSectionLayout
    Dim subName As String

    originalStartRow = this.context.StartRow
    this.context.StartRow = startRow

    this.layout.Configure this.context, startRow, endRow

    sectionName = this.context.Value("main section")
    Set design = this.context.Design

    Set headerRange = this.layout.HeaderRange
    Set printRange = this.layout.PrintHeaderRange
    Set crfRange = this.layout.CrfHeaderRange

    WriteHeader headerRange, sectionName
    WriteHeader printRange, sectionName
    WriteHeader crfRange, sectionName

    this.hooks.FormatSectionHeader this.context, this.layout, headerRange, printRange, crfRange

    boundaryColumns = Array("sub section", "main section", "sheet name")
    currentRow = startRow

    Do While currentRow <= endRow
        this.context.StartRow = currentRow
        subsectionEnd = this.context.EndRow(boundaryColumns)
        If subsectionEnd > endRow Then subsectionEnd = endRow

        Set subLayout = this.hooks.CreateSubSectionLayout
        If subLayout Is Nothing Then
            ThrowError ProjectError.ObjectNotInitialized, _
                       "Section hooks must supply a subsection layout"
        End If

        subLayout.Configure this.context, currentRow, subsectionEnd

        subName = this.context.Value("sub section")

        WriteHeader subLayout.HeaderRange, subName
        WriteHeader subLayout.PrintHeaderRange, subName
        WriteHeader subLayout.CrfHeaderRange, subName

        WriteSubSection subLayout, currentRow, subsectionEnd

        subLayout.Reset
        currentRow = subsectionEnd + 1
    Loop

    this.layout.Reset
    this.context.StartRow = originalStartRow
End Sub

'@section Internal helpers
'===============================================================================
Private Sub WriteHeader(ByVal targetRange As Range, ByVal headerText As String)
    If targetRange Is Nothing Then Exit Sub
    If LenB(headerText) = 0 Then Exit Sub
    targetRange.Cells(1, 1).Value = headerText
End Sub

Private Sub WriteSubSection(ByVal subLayout As ILLSubSectionLayout, _
                             ByVal startRow As Long, _
                             ByVal endRow As Long)

    Dim subWriter As ILLSubSectionWriter
    Dim subHooks As ILLSubSectionWriterHooks

    Set subHooks = this.hooks.SubSectionHooks
    If subHooks Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section hooks must supply subsection hooks"
    End If

    Set subWriter = New LLSubSectionWriterBase
    subWriter.Initialise this.context, subLayout, subHooks
    subWriter.WriteSubSection startRow, endRow
    subWriter.Reset
End Sub
