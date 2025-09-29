VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSubSectionWriterBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Reusable writer orchestrating linelist subsections")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSubSectionWriter

Private Type TSubSectionWriter
    sectionContext As ILLSectionContext
    layout As ILLSubSectionLayout
    hooks As ILLSubSectionWriterHooks
    isReady As Boolean
End Type

Private this As TSubSectionWriter
Private Const CLASSNAME As String = "LLSubSectionWriterBase"

'@section Guard helpers
'===============================================================================
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Sub EnsureReady()
    If Not this.isReady Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Subsection writer has not been initialised"
    End If
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLSubSectionWriter_Initialise(ByVal sectionContext As ILLSectionContext, _
                                           ByVal layout As ILLSubSectionLayout, _
                                           ByVal hooks As ILLSubSectionWriterHooks)

    If sectionContext Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Section context reference is required"
    End If

    If layout Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Subsection layout reference is required"
    End If

    If hooks Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Subsection hooks reference is required"
    End If

    Set this.sectionContext = sectionContext
    Set this.layout = layout
    Set this.hooks = hooks
    this.isReady = True
End Sub

Private Sub ILLSubSectionWriter_Reset()
    Set this.sectionContext = Nothing
    Set this.layout = Nothing
    Set this.hooks = Nothing
    this.isReady = False
End Sub

Private Sub ILLSubSectionWriter_WriteSubSection(ByVal startRow As Long, ByVal endRow As Long)
    Dim originalStartRow As Long
    Dim headerRange As Range
    Dim printRange As Range
    Dim crfRange As Range
    Dim variableNames As BetterArray
    Dim idx As Long
    Dim variableName As String
    Dim varContext As ILLVarContext
    Dim varWriter As ILLVarWriter
    Dim linelist As ILinelist

    EnsureReady

    If startRow <= 0 Or endRow < startRow Then
        ThrowError ProjectError.InvalidArgument, _
                   "Invalid subsection boundaries supplied"
    End If

    originalStartRow = this.sectionContext.StartRow
    this.sectionContext.StartRow = startRow

    this.layout.Configure this.sectionContext, startRow, endRow

    Set headerRange = this.layout.HeaderRange
    Set printRange = this.layout.PrintHeaderRange
    Set crfRange = this.layout.CrfHeaderRange

    this.hooks.FormatHeader this.sectionContext, this.layout, headerRange, printRange, crfRange

    Set variableNames = this.sectionContext.VariableNames(startRow, endRow)
    Set linelist = this.sectionContext.Linelist

    For idx = variableNames.LowerBound To variableNames.UpperBound
        variableName = CStr(variableNames.Item(idx))
        Set varContext = New LLVarContext
        varContext.Initialise linelist, variableName

        Set varWriter = this.hooks.CreateVariableWriter(varContext)
        If varWriter Is Nothing Then
            ThrowError ProjectError.ObjectNotInitialized, _
                       "Subsection hooks must supply a variable writer"
        End If

        varWriter.WriteVariable
        varWriter.Reset
        varContext.ResetCaches
    Next idx

    this.hooks.AfterVariableWrite this.sectionContext, this.layout, headerRange

    this.layout.Reset
    this.sectionContext.StartRow = originalStartRow
End Sub
