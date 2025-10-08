VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSheetNameFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Formats Linelist sheet names consistently across scopes"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Formats Linelist sheet names consistently across scopes")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed, HungarianNotation
'@depends ProjectError

Option Explicit

Implements ILinelistSheetNameFormatter

'@section State and constants
'===============================================================================
Private Type TFormatter
    sheetNameMaxLength As Long                  'Maximum supported length for sheet names including scope prefix
End Type

Private this As TFormatter
Private Const CLASSNAME As String = "LinelistSheetNameFormatter"
Private Const PREFIX_PRINT As String = "print_"
Private Const PREFIX_CRF As String = "crf_"
Private Const DEFAULT_SHEET_MAX_LENGTH As Long = 30


'@section Factory helpers
'===============================================================================
'@sub-title Create a formatter instance with default configuration
Public Function Create() As ILinelistSheetNameFormatter
    Dim formatter As LinelistSheetNameFormatter

    Set formatter = New LinelistSheetNameFormatter
    formatter.Initialise

    Set Create = formatter.Self
End Function

'@sub-title Self reference
Public Property Get Self() As ILinelistSheetNameFormatter
    Set Self = Me
End Property

'@sub-title Initialise internal state
Friend Sub Initialise()
    this.sheetNameMaxLength = DEFAULT_SHEET_MAX_LENGTH
End Sub


'@section Sheet logic
'===============================================================================
'@sub-title Build the scoped sheet name ensuring the length budget is honoured
Public Function FormatSheetName(ByVal baseName As String, _
                                Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As String
    FormatSheetName = FormatSheetNameCore(baseName, scope)
End Function

'@sub-title Determine whether a scoped sheet exists inside a workbook
Public Function SheetExists(ByVal workbook As Workbook, ByVal baseName As String, _
                            Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
    SheetExists = SheetExistsCore(workbook, baseName, scope)
End Function

'@sub-title Retrieve a scoped worksheet reference from the supplied workbook
Public Function ResolveWorksheet(ByVal workbook As Workbook, ByVal baseName As String, _
                                 Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ResolveWorksheet = ResolveWorksheetCore(workbook, baseName, scope)
End Function


'@section Accessors
'===============================================================================
Public Property Get PrintPrefix() As String
    PrintPrefix = PREFIX_PRINT
End Property

Public Property Get CrfPrefix() As String
    CrfPrefix = PREFIX_CRF
End Property

Public Property Get SheetNameMaxLength() As Long
    SheetNameMaxLength = this.sheetNameMaxLength
End Property


'@section Core helpers
'===============================================================================
Private Function FormatSheetNameCore(ByVal baseName As String, _
                                     ByVal scope As LinelistSheetScope) As String
    Dim trimmedName As String
    Dim prefix As String
    Dim candidate As String

    trimmedName = Trim$(baseName)
    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet base name cannot be empty"
    End If

    prefix = ScopePrefix(scope)
    candidate = prefix & trimmedName

    If Len(candidate) > this.sheetNameMaxLength Then
        candidate = Left$(candidate, this.sheetNameMaxLength)
    End If

    FormatSheetNameCore = candidate
End Function

Private Function SheetExistsCore(ByVal workbook As Workbook, _
                                 ByVal baseName As String, _
                                 ByVal scope As LinelistSheetScope) As Boolean
    Dim sheetName As String
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet

    Set targetWorkbook = EnsureWorkbook(workbook)
    sheetName = FormatSheetNameCore(baseName, scope)

    On Error Resume Next
        Set targetSheet = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExistsCore = Not targetSheet Is Nothing
End Function

Private Function ResolveWorksheetCore(ByVal workbook As Workbook, _
                                      ByVal baseName As String, _
                                      ByVal scope As LinelistSheetScope) As Worksheet
    Dim sheetName As String
    Dim targetWorkbook As Workbook

    Set targetWorkbook = EnsureWorkbook(workbook)
    sheetName = FormatSheetNameCore(baseName, scope)

    On Error Resume Next
        Set ResolveWorksheetCore = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ResolveWorksheetCore Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & sheetName & "' not found in workbook '" & targetWorkbook.Name & "'"
    End If
End Function

Private Function EnsureWorkbook(ByVal workbook As Workbook) As Workbook
    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If

    If StrComp(TypeName(workbook), "Workbook", vbTextCompare) <> 0 Then
        ThrowError ProjectError.InvalidArgument, "Workbook argument must be an Excel Workbook"
    End If

    Set EnsureWorkbook = workbook
End Function

Private Function ScopePrefix(ByVal scope As LinelistSheetScope) As String
    Select Case scope
        Case sheetScopeStandard
            ScopePrefix = vbNullString
        Case sheetScopePrint
            ScopePrefix = PREFIX_PRINT
        Case sheetScopeCrf
            ScopePrefix = PREFIX_CRF
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Unknown sheet scope value: " & CStr(scope)
    End Select
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Function ILinelistSheetNameFormatter_FormatSheetName(ByVal baseName As String, _
                                                             Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As String
    ILinelistSheetNameFormatter_FormatSheetName = FormatSheetNameCore(baseName, scope)
End Function

Private Function ILinelistSheetNameFormatter_SheetExists(ByVal workbook As Workbook, ByVal baseName As String, _
                                                         Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
    ILinelistSheetNameFormatter_SheetExists = SheetExistsCore(workbook, baseName, scope)
End Function

Private Function ILinelistSheetNameFormatter_ResolveWorksheet(ByVal workbook As Workbook, ByVal baseName As String, _
                                                              Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
    Set ILinelistSheetNameFormatter_ResolveWorksheet = ResolveWorksheetCore(workbook, baseName, scope)
End Function

Private Property Get ILinelistSheetNameFormatter_PrintPrefix() As String
    ILinelistSheetNameFormatter_PrintPrefix = PREFIX_PRINT
End Property

Private Property Get ILinelistSheetNameFormatter_CrfPrefix() As String
    ILinelistSheetNameFormatter_CrfPrefix = PREFIX_CRF
End Property

Private Property Get ILinelistSheetNameFormatter_SheetNameMaxLength() As Long
    ILinelistSheetNameFormatter_SheetNameMaxLength = this.sheetNameMaxLength
End Property
