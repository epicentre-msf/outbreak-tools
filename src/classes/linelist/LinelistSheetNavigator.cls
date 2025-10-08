VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSheetNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Safely activates linelist worksheets using cached translations and formatted names"

'@Folder("Linelist")
'@ModuleDescription("Safely activates linelist worksheets using cached translations and formatted names")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistSheetNavigator

'@section State and constants
'===============================================================================
Private Type TNavigator
    workbook As Workbook
    formatter As ILinelistSheetNameFormatter
    translations As ITranslationObject
    adminBaseName As String
    instructionBaseName As String
    cacheInitialised As Boolean
    lastActivatedName As String
    translationCacheHits As Long
End Type

Private this As TNavigator
Private Const CLASSNAME As String = "LinelistSheetNavigator"
Private Const ADMIN_KEY As String = "LLSHEET_Admin"
Private Const INSTRUCTION_KEY As String = "INSTSHEETNAME"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal workbook As Workbook, _
                       ByVal translations As ITranslationObject, _
                       Optional ByVal formatter As ILinelistSheetNameFormatter) As ILinelistSheetNavigator

    Dim navigator As LinelistSheetNavigator

    Set navigator = New LinelistSheetNavigator
    navigator.Initialise workbook, translations, formatter

    Set Create = navigator
End Function

Friend Sub Initialise(ByVal workbook As Workbook, _
                      ByVal translations As ITranslationObject, _
                      ByVal formatter As ILinelistSheetNameFormatter)

    Set this.workbook = EnsureWorkbook(workbook)
    Set this.translations = EnsureTranslations(translations)
    Set this.formatter = ResolveFormatter(formatter)

    this.cacheInitialised = False
    this.lastActivatedName = vbNullString
    this.translationCacheHits = 0
End Sub


'@section Public helpers
'===============================================================================
Public Sub ActivateAdminSheet()
    EnsureCoreNames
    ActivateAndTrack this.adminBaseName, sheetScopeStandard
End Sub

Public Sub ActivateInstructionSheet()
    EnsureCoreNames
    ActivateAndTrack this.instructionBaseName, sheetScopeStandard
End Sub

Public Function ActivateSheet(ByVal baseName As String, _
                              Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ActivateSheet = ActivateAndTrack(baseName, scope)
End Function

Public Function ResolveSheet(ByVal baseName As String, _
                             Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ResolveSheet = ResolveSheetCore(baseName, scope)
End Function

Public Function SheetExists(ByVal baseName As String, _
                            Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
    Dim trimmedName As String

    trimmedName = NormaliseName(baseName)
    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet base name cannot be empty"
    End If

    SheetExists = this.formatter.SheetExists(this.workbook, trimmedName, scope)
End Function

Public Property Get LastActivatedSheetName() As String
    LastActivatedSheetName = this.lastActivatedName
End Property

Public Property Get TranslationCacheHits() As Long
    TranslationCacheHits = this.translationCacheHits
End Property


'@section Core helpers
'===============================================================================
Private Function ActivateAndTrack(ByVal baseName As String, _
                                  ByVal scope As LinelistSheetScope) As Worksheet
    Dim sheet As Worksheet

    Set sheet = ResolveSheetCore(baseName, scope)

    On Error GoTo ActivateFail
        sheet.Activate
        this.lastActivatedName = sheet.Name
        Set ActivateAndTrack = sheet
        On Error GoTo 0
        Exit Function

ActivateFail:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Unable to activate worksheet '" & sheet.Name & "': " & Err.Description
End Function

Private Function ResolveSheetCore(ByVal baseName As String, _
                                  ByVal scope As LinelistSheetScope) As Worksheet
    Dim trimmedName As String

    trimmedName = NormaliseName(baseName)
    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet base name cannot be empty"
    End If

    Set ResolveSheetCore = this.formatter.ResolveWorksheet(this.workbook, trimmedName, scope)
End Function

Private Sub EnsureCoreNames()
    If this.cacheInitialised Then Exit Sub

    this.adminBaseName = LookupTranslation(ADMIN_KEY)
    this.instructionBaseName = LookupTranslation(INSTRUCTION_KEY)

    this.cacheInitialised = True
End Sub

Private Function LookupTranslation(ByVal tagName As String) As String
    Dim translated As String

    translated = Trim$(CStr(this.translations.TranslatedValue(tagName)))

    this.translationCacheHits = this.translationCacheHits + 1

    If LenB(translated) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Translation key '" & tagName & "' returned an empty value"
    End If

    LookupTranslation = translated
End Function

Private Function NormaliseName(ByVal baseName As String) As String
    NormaliseName = Trim$(baseName)
End Function

Private Function EnsureWorkbook(ByVal workbook As Workbook) As Workbook
    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If

    If StrComp(TypeName(workbook), "Workbook", vbTextCompare) <> 0 Then
        ThrowError ProjectError.InvalidArgument, "Navigator expects an Excel workbook"
    End If

    Set EnsureWorkbook = workbook
End Function

Private Function EnsureTranslations(ByVal translations As ITranslationObject) As ITranslationObject
    If translations Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object cannot be Nothing"
    End If

    Set EnsureTranslations = translations
End Function

Private Function ResolveFormatter(ByVal formatter As ILinelistSheetNameFormatter) As ILinelistSheetNameFormatter
    If formatter Is Nothing Then
        Set ResolveFormatter = LinelistSheetNameFormatter.Create
    Else
        Set ResolveFormatter = formatter
    End If
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistSheetNavigator_ActivateAdminSheet()
    ActivateAdminSheet
End Sub

Private Sub ILinelistSheetNavigator_ActivateInstructionSheet()
    ActivateInstructionSheet
End Sub

Private Function ILinelistSheetNavigator_ActivateSheet(ByVal baseName As String, _
                                                       Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ILinelistSheetNavigator_ActivateSheet = ActivateSheet(baseName, scope)
End Function

Private Function ILinelistSheetNavigator_ResolveSheet(ByVal baseName As String, _
                                                      Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ILinelistSheetNavigator_ResolveSheet = ResolveSheet(baseName, scope)
End Function

Private Function ILinelistSheetNavigator_SheetExists(ByVal baseName As String, _
                                                     Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
    ILinelistSheetNavigator_SheetExists = SheetExists(baseName, scope)
End Function

Private Property Get ILinelistSheetNavigator_LastActivatedSheetName() As String
    ILinelistSheetNavigator_LastActivatedSheetName = LastActivatedSheetName
End Property

Private Property Get ILinelistSheetNavigator_TranslationCacheHits() As Long
    ILinelistSheetNavigator_TranslationCacheHits = TranslationCacheHits
End Property

