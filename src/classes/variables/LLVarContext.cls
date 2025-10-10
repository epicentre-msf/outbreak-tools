VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Shared context cache for Linelist variable writers")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarContext

Private Type TLLVarContext
    ll As ILinelist
    varName As String
    dictionary As ILLdictionary
    variables As ILLVariables
    specs As ILinelistSpecs
    dropdownDefault As IDropdownLists
    dropdownCustom As IDropdownLists
    dropdownCustomScope As Byte
    cachedSheet As Worksheet
    cachedSheetName As String
    cachedPrintSheet As Worksheet
    cachedPrintSheetName As String
    hasLL As Boolean
End Type

Private this As TLLVarContext
Private Const CLASSNAME As String = "LLVarContext"
Private Const COL_SHEET_NAME As String = "sheet name"

'@section Factory
'===============================================================================
'@description Helper factory mirroring the pattern used across linelist classes.
Public Function Create(ByVal ll As ILinelist, ByVal varName As String) As ILLVarContext
    With New LLVarContext
        .Initialise ll, varName
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLVarContext
    Set Self = Me
End Property

'@section Cache management
'===============================================================================
Private Sub ResetSheetCaches()
    Set this.cachedSheet = Nothing
    this.cachedSheetName = vbNullString
    Set this.cachedPrintSheet = Nothing
    this.cachedPrintSheetName = vbNullString
End Sub

Private Sub ResetCollaboratorCaches()
    Set this.dropdownDefault = Nothing
    Set this.dropdownCustom = Nothing
    this.dropdownCustomScope = 0
End Sub

'@section Guards
'===============================================================================
Private Sub EnsureInitialised()
    If Not this.hasLL Then
        If this.ll Is Nothing Then
            ThrowError ProjectError.ObjectNotInitialized, "Linelist Object has not been initialised"
        End If
        If LenB(this.varName) = 0 Then
            ThrowError ProjectError.ObjectNotInitialized, "Variable name is not set on the context"
        End If
        this.hasLL = True
    End If
End Sub

Private Sub EnsureDictionary()
    EnsureInitialised
    If this.dictionary Is Nothing Then
        Set this.dictionary = this.ll.Dictionary()
    End If
    If this.dictionary Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist dictionary reference is missing"
    End If
End Sub

Private Sub EnsureSpecs()
    EnsureInitialised
    If this.specs Is Nothing Then
        Set this.specs = this.ll.LinelistData()
    End If
    If this.specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist specifications reference is missing"
    End If
End Sub

Private Sub EnsureVariables()
    EnsureDictionary
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dictionary)
    End If
End Sub

Private Function NormaliseColumnName(ByVal columnName As String) As String
    NormaliseColumnName = Trim$(columnName)
End Function

Private Function NormaliseVariableName(ByVal varName As String) As String
    NormaliseVariableName = Trim$(varName)
End Function

Private Sub ThrowError(ByVal errNumber As Byte, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, errorMessage
End Sub

'@section Worksheet helpers
'===============================================================================
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet

    If LenB(sheetName) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
            "Unable to resolve linelist worksheet because the sheet name column is empty"
    End If

    On Error Resume Next
        Set target = this.ll.Wksh(sheetName)
    On Error GoTo 0

    If target Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
            "Worksheet '" & sheetName & "' could not be located for the linelist variable"
    End If

    Set ResolveWorksheet = target
End Function

Private Function ResolvePrintWorksheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet

    If LenB(sheetName) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
            "Unable to resolve printed worksheet because the sheet name column is empty"
    End If

    On Error Resume Next
        Set target = this.ll.PrintWksh(sheetName)
    On Error GoTo 0

    If target Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
            "Printed worksheet '" & sheetName & "' could not be located for the linelist variable"
    End If

    Set ResolvePrintWorksheet = target
End Function

Private Function ResolveDropdownInternal(ByVal scope As Byte) As IDropdownLists
    Dim drop As IDropdownLists

    On Error Resume Next
        Set drop = this.ll.Dropdown(scope:=scope)
    On Error GoTo 0

    If drop Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
            "Dropdown scope '" & CStr(scope) & "' is not available on the linelist"
    End If

    Set ResolveDropdownInternal = drop
End Function

'@section Interface implementation
'===============================================================================
Private Sub ILLVarContext_Initialise(ByVal ll As ILinelist, _
                                     ByVal varName As String, _
                                     Optional ByVal dictionary As ILLdictionary, _
                                     Optional ByVal variables As ILLVariables, _
                                     Optional ByVal specs As ILinelistSpecs)

    Dim normalisedName As String

    If ll Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A linelist reference is required to initialise the context"
    End If

    normalisedName = NormaliseVariableName(varName)
    If LenB(normalisedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Variable name is required to initialise the context"
    End If

    Set this.ll = ll
    this.varName = normalisedName

    Set this.dictionary = dictionary
    Set this.variables = variables
    Set this.specs = specs

    ResetSheetCaches
    ResetCollaboratorCaches
End Sub

Private Property Get ILLVarContext_Linelist() As ILinelist
    EnsureInitialised
    Set ILLVarContext_Linelist = this.ll
End Property

Private Property Get ILLVarContext_VariableName() As String
    EnsureInitialised
    ILLVarContext_VariableName = this.varName
End Property

Private Sub ILLVarContext_ResetCaches()
    ResetSheetCaches
    ResetCollaboratorCaches
    Set this.variables = Nothing
    Set this.dictionary = Nothing
    Set this.specs = Nothing
End Sub

Private Function ILLVarContext_ValueOf(ByVal columnName As String) As String
    Dim columnKey As String
    EnsureVariables

    columnKey = NormaliseColumnName(columnName)
    If LenB(columnKey) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Column name must not be empty when requesting metadata"
    End If

    ILLVarContext_ValueOf = this.variables.Value(colName:=columnKey, varName:=this.varName)
End Function

Private Function ILLVarContext_Dictionary() As ILLdictionary
    EnsureDictionary
    Set ILLVarContext_Dictionary = this.dictionary
End Function

Private Function ILLVarContext_Variables() As ILLVariables
    EnsureVariables
    Set ILLVarContext_Variables = this.variables
End Function

Private Function ILLVarContext_Specs() As ILinelistSpecs
    EnsureSpecs
    Set ILLVarContext_Specs = this.specs
End Function

Private Function ILLVarContext_Design() As ILLFormat
    EnsureSpecs
    Set ILLVarContext_Design = this.specs.DesignFormat
End Function

Private Function ILLVarContext_Translation() As ITranslationObject
    EnsureSpecs
    Set ILLVarContext_Translation = this.specs.TransObject
End Function

Private Function ILLVarContext_Categories(ByVal varName As String, _
                                          Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim requestedName As String
    requestedName = NormaliseVariableName(IIf(LenB(varName) = 0, this.varName, varName))
    EnsureSpecs
    Set ILLVarContext_Categories = this.specs.Categories(requestedName, useShortlabels)
End Function

Private Function ILLVarContext_EditableLabelFlag() As String
    ILLVarContext_EditableLabelFlag = LCase$(Trim$(ILLVarContext_ValueOf("editable label")))
End Function

Private Function ILLVarContext_RegisterBookDirective() As String
    ILLVarContext_RegisterBookDirective = LCase$(Trim$(ILLVarContext_ValueOf("register book")))
End Function

Private Function ILLVarContext_FormattingConditionName() As String
    ILLVarContext_FormattingConditionName = Trim$(ILLVarContext_ValueOf("formatting condition"))
End Function

Private Function ILLVarContext_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    EnsureInitialised

    Select Case scope
    Case 1
        If this.dropdownDefault Is Nothing Then
            Set this.dropdownDefault = ResolveDropdownInternal(scope)
        End If
        Set ILLVarContext_Dropdown = this.dropdownDefault

    Case Else
        If (this.dropdownCustom Is Nothing) Or (this.dropdownCustomScope <> scope) Then
            Set this.dropdownCustom = ResolveDropdownInternal(scope)
            this.dropdownCustomScope = scope
        End If
        Set ILLVarContext_Dropdown = this.dropdownCustom
    End Select
End Function

Private Function ILLVarContext_Worksheet() As Worksheet
    Dim sheetName As String

    EnsureVariables

    sheetName = ILLVarContext_ValueOf(COL_SHEET_NAME)

    If (this.cachedSheet Is Nothing) Or (StrComp(this.cachedSheetName, sheetName, vbTextCompare) <> 0) Then
        Set this.cachedSheet = ResolveWorksheet(sheetName)
        this.cachedSheetName = sheetName
    End If

    Set ILLVarContext_Worksheet = this.cachedSheet
End Function

Private Function ILLVarContext_PrintWorksheet() As Worksheet
    Dim sheetName As String

    EnsureVariables

    sheetName = ILLVarContext_ValueOf(COL_SHEET_NAME)

    If (this.cachedPrintSheet Is Nothing) Or _
       (StrComp(this.cachedPrintSheetName, sheetName, vbTextCompare) <> 0) Then
        Set this.cachedPrintSheet = ResolvePrintWorksheet(sheetName)
        this.cachedPrintSheetName = sheetName
    End If

    Set ILLVarContext_PrintWorksheet = this.cachedPrintSheet
End Function
