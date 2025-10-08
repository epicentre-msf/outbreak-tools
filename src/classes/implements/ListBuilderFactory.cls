VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuilderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Creates fully wired list builder coordinators based on worksheet metadata")
'@depends ProjectError, LLSheets, ListBuilderCoordinator, DictionarySectionBuilder, ListWorksheetPreparer, ListBuildContext
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TFactoryState
    contextCache As IListContextCache
End Type

Private this As TFactoryState
Private Const CLASSNAME As String = "ListBuilderFactory"
Private Const SHEETTYPE_HORIZONTAL As String = "hlist2d"
Private Const SHEETTYPE_VERTICAL As String = "vlist1d"
Private Const TABLETYPE_HORIZONTAL As Byte = 1
Private Const TABLETYPE_VERTICAL As Byte = 2

'@section Factory
'===============================================================================
'Create a factory optionally using the supplied context cache.
Public Function Create(Optional ByVal cache As IListContextCache = Nothing) As ListBuilderFactory
    With New ListBuilderFactory
        If cache Is Nothing Then
            Set .ContextCache = ListContextCache.Create
        Else
            Set .ContextCache = cache
        End If
        Set Create = .Self
    End With
End Function

'Expose the instance for fluent chaining.
Public Property Get Self() As ListBuilderFactory
    Set Self = Me
End Property

'Assign the context cache after validation.
Public Property Set ContextCache(ByVal cache As IListContextCache)
    If cache Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "contextCache cannot be Nothing"
    End If
    Set this.contextCache = cache
End Property

'@section Coordinator creation
'===============================================================================
'Create a coordinator ready to build the requested worksheet.
Public Function CreateCoordinator(ByVal context As IListBuildContext) As IListBuilderCoordinator
    ValidateContext context

    Dim linelist As ILinelist
    Set linelist = TryCastLinelist(context.Linelist)
    this.contextCache.AttachLinelist linelist

    Dim dictionary As ILLdictionary
    Set dictionary = linelist.Dictionary
    ValidateDictionary dictionary
    this.contextCache.AttachDictionary dictionary

    Dim layout As IListLayoutStrategy
    Set layout = ResolveLayoutStrategy(dictionary, context.WorksheetName)

    Dim sectionBuilder As IListSectionBuilder
    Set sectionBuilder = DictionarySectionBuilder.Create(this.contextCache)

    Dim preparer As IListWorksheetPreparer
    Set preparer = New ListWorksheetPreparer

    Set CreateCoordinator = ListBuilderCoordinator.Create(layout, sectionBuilder, preparer)
End Function

'Create a build context for the supplied worksheet and linelist.
Public Function CreateContext(ByVal linelist As ILinelist, _
                              ByVal worksheetName As String) As IListBuildContext

    ValidateWorksheetName worksheetName
    If linelist Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "linelist cannot be Nothing when creating a context"
    End If

    Dim dictionary As ILLdictionary
    Set dictionary = linelist.Dictionary
    ValidateDictionary dictionary

    Dim sheetHelper As ILLSheets
    Set sheetHelper = LLSheets.Create(dictionary)

    Dim sheetType As String
    sheetType = NormalizeValue(sheetHelper.SheetInfo(worksheetName, SheetInfoType.SheetInfoSheetType))

    Dim tableTypeCode As Byte
    tableTypeCode = TableTypeFromSheetType(sheetType)

    this.contextCache.AttachLinelist linelist
    this.contextCache.AttachDictionary dictionary

    Set CreateContext = ListBuildContext.Create(worksheetName, tableTypeCode, linelist)
End Function

'@section Helpers
'===============================================================================
Private Function ResolveLayoutStrategy(ByVal dictionary As ILLdictionary, _
                                       ByVal sheetName As String) As IListLayoutStrategy

    Dim sheetHelper As ILLSheets
    Set sheetHelper = LLSheets.Create(dictionary)

    Dim sheetType As String
    sheetType = NormalizeValue(sheetHelper.SheetInfo(sheetName, SheetInfoType.SheetInfoSheetType))

    Select Case sheetType
        Case SHEETTYPE_HORIZONTAL
            Set ResolveLayoutStrategy = New HorizontalListLayoutStrategy
        Case Else
            Set ResolveLayoutStrategy = New VerticalListLayoutStrategy
    End Select
End Function

Private Function TableTypeFromSheetType(ByVal sheetType As String) As Byte
    Select Case sheetType
        Case SHEETTYPE_HORIZONTAL
            TableTypeFromSheetType = TABLETYPE_HORIZONTAL
        Case Else
            TableTypeFromSheetType = TABLETYPE_VERTICAL
    End Select
End Function

Private Function NormalizeValue(ByVal valueText As String) As String
    NormalizeValue = LCase$(Trim$(valueText))
End Function

Private Sub ValidateWorksheetName(ByVal worksheetName As String)
    If Trim$(worksheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Worksheet name cannot be empty"
    End If
End Sub

Private Sub ValidateContext(ByVal context As IListBuildContext)
    If context Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Build context cannot be Nothing"
    End If

    ValidateWorksheetName context.WorksheetName
End Sub

Private Sub ValidateDictionary(ByVal dictionary As ILLdictionary)
    If dictionary Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Dictionary reference is missing"
    End If
End Sub

Private Function TryCastLinelist(ByVal candidate As Object) As ILinelist
    If candidate Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Linelist reference is required"
    End If

    If NormalizeValue(TypeName(candidate)) <> NormalizeValue("ILinelist") Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Linelist reference must implement ILinelist"
    End If

    Set TryCastLinelist = candidate
End Function

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

