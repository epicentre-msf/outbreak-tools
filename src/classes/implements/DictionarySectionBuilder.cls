VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionarySectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Builds list sections using dictionary metadata via LLSheets")
'@depends ProjectError, LLSheets, ListSectionDescriptor
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListSectionBuilder

Private Type TBuilderState
    contextCache As IListContextCache
End Type

Private this As TBuilderState
Private Const CLASSNAME As String = "DictionarySectionBuilder"

'@section Factory
'===============================================================================
'Create a section builder optionally primed with a context cache.
Public Function Create(Optional ByVal cache As IListContextCache = Nothing) As IListSectionBuilder
    With New DictionarySectionBuilder
        If cache Is Nothing Then
            Set .ContextCache = ListContextCache.Create
        Else
            Set .ContextCache = cache
        End If
        Set Create = .Self
    End With
End Function

'Expose the instance as the interface type.
Public Property Get Self() As IListSectionBuilder
    Set Self = Me
End Property

'Internal setter used by the factory.
Friend Property Set ContextCache(ByVal cache As IListContextCache)
    If cache Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "Context cache cannot be Nothing"
    End If
    Set this.contextCache = cache
End Property

'@section Interface implementation
'===============================================================================
Private Function IListSectionBuilder_BuildSections(ByVal context As IListBuildContext) As BetterArray
    ValidateContext context

    Dim sections As BetterArray
    Set sections = New BetterArray
    sections.LowerBound = 1

    Dim linelist As Object
    Set linelist = context.Linelist
    If linelist Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Linelist reference is required to build sections"
    End If

    this.contextCache.AttachLinelist linelist

    Dim dictionary As ILLdictionary
    Set dictionary = linelist.Dictionary
    this.contextCache.AttachDictionary dictionary

    Dim sheetHelper As ILLSheets
    Set sheetHelper = LLSheets.Create(dictionary)

    Dim sheetName As String
    sheetName = context.WorksheetName

    If Not sheetHelper.Contains(sheetName) Then
        ThrowError CLng(ProjectError.ElementNotFound), _
                  "Worksheet '" & sheetName & "' is not registered in the dictionary"
    End If

    Dim anchorRow As Long
    Dim anchorColumn As Long

    anchorRow = sheetHelper.DataBounds(sheetName, SheetBound.RowSart)
    anchorColumn = sheetHelper.DataBounds(sheetName, SheetBound.ColStart)

    If anchorRow <= 0 Then anchorRow = sheetHelper.RowIndex(sheetName)
    If anchorRow <= 0 Then anchorRow = 1
    If anchorColumn <= 0 Then anchorColumn = 1

    sections.Push ListSectionDescriptor.Create(sheetName & "__main", anchorRow, anchorColumn)

    Set IListSectionBuilder_BuildSections = sections
End Function

Private Sub IListSectionBuilder_Invalidate()
    On Error Resume Next
        this.contextCache.Invalidate
    On Error GoTo 0
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateContext(ByVal context As IListBuildContext)
    If context Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Build context cannot be Nothing"
    End If

    If Trim$(context.WorksheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Worksheet name is required when building dictionary sections"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

