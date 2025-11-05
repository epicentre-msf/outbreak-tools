VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterSetupPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation for Master Setup preparation and configuration"

'@PredeclaredId
'@Folder("MasterSetup")
'@ModuleDescription("Implementation for Master Setup preparation and configuration")
'@depends ApplicationState, IApplicationState, DropdownLists, IDropdownLists, BetterArray, MasterSetupVariables, IMasterSetupVariables, ProjectError
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IMasterSetupPreparation

'@section State
'===============================================================================
Private Type TMasterSetupPreparation
    hostBook As Workbook             'Workbook hosting master setup artefacts
    dropdownSheet As Worksheet       'Worksheet hosting dropdown listobjects
    variablesSheet As Worksheet      'Worksheet hosting the variables table
    dropdowns As IDropdownLists      'Cached dropdown manager
    variables As IMasterSetupVariables  'Cached variables manager
End Type

Private this As TMasterSetupPreparation

'@section Constants
'===============================================================================
Private Const CLASS_NAME As String = "MasterSetupPreparation"
Private Const DROPDOWN_SHEET_NAME As String = "__dropdowns"
Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const DEFAULT_VARIABLE_TABLE_NAME As String = "Tab_MasterVariables"
Private Const DROPDOWN_PREFIX As String = "dropdown_"
Private Const LANGUAGES_DROPDOWN As String = "__languages"
Private Const FILE_LANGUAGES_DROPDOWN As String = "__file_languages"
Private Const DISEASES_DROPDOWN_PROHIBITED As String = "__prohibited_diseases_list"
Private Const DISEASES_DROPODWN As String  = "__diseases_list"
Private Const CONFIGS_DROPDOWN As String = "__configSheets"
Private Const STATUS_DROPDOWN As String = "__var_status"
Private Const YESNO_DROPDOWN As String = "__yesno"
Private Const STATUS_ACTIVE As String = "active"
Private Const STATUS_INACTIVE As String = "inactive"
Private Const VARIABLE_TABLE_STYLE As String = "TableStyleMedium2"


'@section Factory helpers
'===============================================================================
'@Description("Create a configured MasterSetupPreparation instance for the supplied workbook.")
Public Function Create(ByVal hostBook As Workbook) As IMasterSetupPreparation
    Dim instance As MasterSetupPreparation

    Set instance = New MasterSetupPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@Description("Return the current instance as an interface reference.")
Public Property Get Self() As IMasterSetupPreparation
    Set Self = Me
End Property


'@section Public API
'===============================================================================
'@Description("Bind the preparation helper to the workbook hosting master setup assets.")
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required"
    End If

    Set this.hostBook = hostBook
    ResetCaches
End Sub

'@Description("Run the preparation workflow: register dropdowns then initialise the variables table.")
Public Sub Prepare(Optional ByVal excelApplication As Application)
    Dim scope As IApplicationState
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo PrepareFail

    Set scope = ApplicationState.Create(excelApplication)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, busyCursor:=xlWait

    EnsureDropdowns
    EnsureVariables

    On Error GoTo 0

Cleanup:
    On Error Resume Next
        If Not scope Is Nothing Then
            scope.Restore
        End If
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Sub

PrepareFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume Cleanup
End Sub

'@Description("Ensure the dropdown worksheet hosts all required dropdown lists.")
Public Sub EnsureDropdowns()
    Dim drop As IDropdownLists
    Dim definitions As Variant
    Dim idx As Long
    Dim entry As Variant

    Set drop = EnsureDropdownManager
    definitions = DefaultDropdownDefinitions()

    For idx = LBound(definitions) To UBound(definitions)
        entry = definitions(idx)
        EnsureDropdownDefinition drop, CStr(entry(0)), entry(1)
    Next idx

    RefreshLanguageDropdowns drop
End Sub

'@Description("Ensure the variables worksheet exposes an initialised manager with validations applied.")
Public Sub EnsureVariables()
    Dim manager As IMasterSetupVariables

    Set manager = EnsureVariablesManager
    If manager Is Nothing Then Exit Sub

    If Not manager.Initialised Then
        manager.Initialise Dropdowns
    End If
End Sub

'@Description("Expose the dropdown manager for callers needing direct access.")
Public Property Get Dropdowns() As IDropdownLists
    EnsureDropdownManager
    Set Dropdowns = this.dropdowns
End Property

'@Description("Expose the variables manager associated with the master setup table.")
Public Property Get Variables() As IMasterSetupVariables
    EnsureVariablesManager
    Set Variables = this.variables
End Property

'@Description("Expose the configured host workbook.")
Public Property Get HostWorkbook() As Workbook
    Set HostWorkbook = EnsureWorkbook
End Property


'@section Preparation helpers
'===============================================================================
'@Description("Reset cached dependencies after reconfiguration.")
Private Sub ResetCaches()
    Set this.dropdownSheet = Nothing
    Set this.variablesSheet = Nothing
    Set this.dropdowns = Nothing
    Set this.variables = Nothing
End Sub

'@Description("Ensure the dropdown manager is initialised and cached.")
Private Function EnsureDropdownManager() As IDropdownLists
    If Not this.dropdowns Is Nothing Then
        Set EnsureDropdownManager = this.dropdowns
        Exit Function
    End If

    Dim wsDrop As Worksheet

    Set wsDrop = EnsureDropdownSheet
    Set this.dropdowns = DropdownLists.Create(wsDrop, DROPDOWN_PREFIX)
    Set EnsureDropdownManager = this.dropdowns
End Function

'@Description("Ensure the variables manager is initialised over an existing listobject.")
Private Function EnsureVariablesManager() As IMasterSetupVariables
    If Not this.variables Is Nothing Then
        Set EnsureVariablesManager = this.variables
        Exit Function
    End If

    Dim targetTable As ListObject

    Set targetTable = EnsureVariablesTable
    If targetTable Is Nothing Then Exit Function

    Set this.variables = MasterSetupVariables.Create(targetTable)
    Set EnsureVariablesManager = this.variables
End Function

'@Description("Ensure the dropdown worksheet exists and cache the reference.")
Private Function EnsureDropdownSheet() As Worksheet
    If Not this.dropdownSheet Is Nothing Then
        Set EnsureDropdownSheet = this.dropdownSheet
        Exit Function
    End If

    Dim wsDrop As Worksheet

    Set wsDrop = ResolveOrCreateWorksheet(DROPDOWN_SHEET_NAME, xlSheetVisible)
    Set this.dropdownSheet = wsDrop
    Set EnsureDropdownSheet = wsDrop
End Function

'@Description("Ensure the variables worksheet exists and cache the reference.")
Private Function EnsureVariablesSheet() As Worksheet
    If Not this.variablesSheet Is Nothing Then
        Set EnsureVariablesSheet = this.variablesSheet
        Exit Function
    End If

    Dim wsVars As Worksheet

    Set wsVars = ResolveOrCreateWorksheet(VARIABLES_SHEET_NAME, xlSheetVisible)
    Set this.variablesSheet = wsVars
    Set EnsureVariablesSheet = wsVars
End Function

'@Description("Ensure a variables table exists on the variables worksheet.")
Private Function EnsureVariablesTable() As ListObject
    Dim wsVars As Worksheet
    Dim table As ListObject

    Set wsVars = EnsureVariablesSheet
    Set table = ResolveVariablesTable(wsVars)

    If table Is Nothing Then
        Set table = CreateVariablesTable(wsVars)
    End If

    EnsureTableHasRow table
    ApplyTableStyle table, VARIABLE_TABLE_STYLE

    Set EnsureVariablesTable = table
End Function

'@Description("Resolve the primary variables table if it already exists.")
Private Function ResolveVariablesTable(ByVal wsVars As Worksheet) As ListObject
    Dim table As ListObject

    If wsVars Is Nothing Then Exit Function

    On Error Resume Next
        Set table = wsVars.ListObjects(DEFAULT_VARIABLE_TABLE_NAME)
    On Error GoTo 0

    If table Is Nothing Then
        If wsVars.ListObjects.Count > 0 Then
            Set table = wsVars.ListObjects(1)
        End If
    End If

    Set ResolveVariablesTable = table
End Function

'@Description("Create a variables table with default headers when none exists.")
Private Function CreateVariablesTable(ByVal wsVars As Worksheet) As ListObject
    Dim headers As Variant
    Dim headerCount As Long
    Dim headerRange As Range
    Dim dataRange As Range
    Dim idx As Long
    Dim table As ListObject

    If wsVars Is Nothing Then Exit Function

    headers = DefaultVariableHeaders()
    headerCount = UBound(headers) - LBound(headers) + 1

    Set headerRange = wsVars.Cells(1, 1).Resize(1, headerCount)
    For idx = 0 To headerCount - 1
        headerRange.Cells(1, idx + 1).Value = headers(idx)
    Next idx

    Set dataRange = headerRange.Resize(2, headerCount)
    dataRange.Rows(2).ClearContents

    Set table = wsVars.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    table.Name = DEFAULT_VARIABLE_TABLE_NAME

    Set CreateVariablesTable = table
End Function

'@Description("Ensure the supplied table exposes at least one data row.")
Private Sub EnsureTableHasRow(ByVal table As ListObject)
    Dim dataRange As Range

    If table Is Nothing Then Exit Sub

    On Error Resume Next
        Set dataRange = table.DataBodyRange
    On Error GoTo 0

    If dataRange Is Nothing Then
        table.ListRows.Add
        Exit Sub
    End If

    If dataRange.Rows.Count = 0 Then
        table.ListRows.Add
    End If
End Sub

'@Description("Apply a table style while ignoring styling errors on unsupported versions.")
Private Sub ApplyTableStyle(ByVal table As ListObject, ByVal tableStyle As String)
    If table Is Nothing Then Exit Sub
    If LenB(tableStyle) = 0 Then Exit Sub

    On Error Resume Next
        table.TableStyle = tableStyle
    On Error GoTo 0
End Sub

'@Description("Return the default header labels for a freshly created variables table.")
Private Function DefaultVariableHeaders() As Variant
    DefaultVariableHeaders = Array("Variable Order", "Variable Section", "Variable Name", _
                                   "Label", "Default Choice", "Choices Values", _
                                   "Default Status", "Comments")
End Function

'@Description("Supply the default dropdown definitions required by master setup.")
Private Function DefaultDropdownDefinitions() As Variant
    DefaultDropdownDefinitions = Array( _
        Array(YESNO_DROPDOWN, Array("yes", "no")), _
        Array(STATUS_DROPDOWN, Array(STATUS_ACTIVE, STATUS_INACTIVE)), _
        Array(DISEASES_DROPDOWN_PROHIBITED, DefaultDiseaseExclusions()), _
        Array(CONFIGS_DROPDOWN, DefaultConfigSheets()), _
        Array(DISEASES_DROPODWN, Array("", "")), _
        Array(FILE_LANGUAGES_DROPDOWN, Array(vbNullString)), _
        Array(LANGUAGES_DROPDOWN, Array(vbNullString)) _
    )
End Function

'@Description("Return the default list of worksheets excluded from disease dropdowns.")
Private Function DefaultDiseaseExclusions() As Variant
    DefaultDiseaseExclusions = Array("Choices", "Variables", "Translations", "Dev", _
                                     "__pass", "__updated", "__compRep", "__ribbonTranslation", _
                                     "__dropdowns")
End Function

'@Description("Return the default configuration worksheets list.")
Private Function DefaultConfigSheets() As Variant
    DefaultConfigSheets = Array("Dev", "__pass", "__updated", "__compRep", "__ribbonTranslation", "__dropdowns")
End Function

'@Description("Ensure a dropdown definition exists with the supplied values.")
Private Sub EnsureDropdownDefinition(ByVal drop As IDropdownLists, _
                                     ByVal listName As String, _
                                     ByVal items As Variant)
    Dim values As BetterArray

    Set values = BuildArray(items)
    EnsureDropdownHasValues drop, listName, values
End Sub

'@Description("Populate the languages dropdowns using the translations sheet headers.")
Private Sub RefreshLanguageDropdowns(ByVal drop As IDropdownLists)
    Dim wsTrans As Worksheet
    Dim translations As ListObject
    Dim langValues As BetterArray

    If drop Is Nothing Then Exit Sub

    Set wsTrans = ResolveWorksheet(TRANSLATIONS_SHEET_NAME)
    If wsTrans Is Nothing Then Exit Sub
    If wsTrans.ListObjects.Count = 0 Then Exit Sub

    Set translations = wsTrans.ListObjects(1)
    If translations.HeaderRowRange Is Nothing Then Exit Sub

    Set langValues = New BetterArray
    langValues.LowerBound = 1
    langValues.FromExcelRange translations.HeaderRowRange

    If langValues.Length = 0 Then
        langValues.Push vbNullString
    End If

    EnsureDropdownHasValues drop, LANGUAGES_DROPDOWN, langValues
    EnsureDropdownHasValues drop, FILE_LANGUAGES_DROPDOWN, langValues.Clone
End Sub

'@Description("Convert supplied items into a BetterArray helper.")
Private Function BuildArray(ByVal items As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.LowerBound = 1

    If IsArray(items) Then
        For idx = LBound(items) To UBound(items)
            result.Push items(idx)
        Next idx
    Else
        result.Push items
    End If

    Set BuildArray = result
End Function

'@Description("Ensure a dropdown either exists or is updated with the supplied values.")
Private Sub EnsureDropdownHasValues(ByVal drop As IDropdownLists, _
                                    ByVal listName As String, _
                                    ByVal values As BetterArray)
    If drop Is Nothing Then Exit Sub
    If values Is Nothing Then Exit Sub

    If values.Length = 0 Then
        values.Push vbNullString
    End If

    If drop.Exists(listName) Then
        drop.Update values, listName, removeDuplicates:=True
    Else
        drop.Add values, listName
    End If
End Sub

'@Description("Resolve a worksheet by name without creating it.")
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim wbHost As Workbook

    Set wbHost = EnsureWorkbook

    On Error Resume Next
        Set ResolveWorksheet = wbHost.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@Description("Resolve a worksheet or create it when absent.")
Private Function ResolveOrCreateWorksheet(ByVal sheetName As String, _
                                          Optional ByVal visibility As XlSheetVisibility = xlSheetVisible, _
                                          Optional ByVal applyVisibility As Boolean = True) As Worksheet
    Dim wbHost As Workbook
    Dim wsTarget As Worksheet

    Set wbHost = EnsureWorkbook

    On Error Resume Next
        Set wsTarget = wbHost.Worksheets(sheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Set wsTarget = wbHost.Worksheets.Add(After:=wbHost.Worksheets(wbHost.Worksheets.Count))
        wsTarget.Name = sheetName
    End If

    If applyVisibility Then
        wsTarget.Visible = visibility
    End If

    Set ResolveOrCreateWorksheet = wsTarget
End Function

'@Description("Ensure the host workbook reference is available.")
Private Function EnsureWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   CLASS_NAME & " has not been configured with a workbook reference"
    End If

    Set EnsureWorkbook = this.hostBook
End Function


'@section Interface implementation
'===============================================================================
Private Sub IMasterSetupPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Sub IMasterSetupPreparation_Prepare(Optional ByVal excelApplication As Application)
    Prepare excelApplication
End Sub

Private Sub IMasterSetupPreparation_EnsureDropdowns()
    EnsureDropdowns
End Sub

Private Sub IMasterSetupPreparation_EnsureVariables()
    EnsureVariables
End Sub

Private Property Get IMasterSetupPreparation_Dropdowns() As IDropdownLists
    Set IMasterSetupPreparation_Dropdowns = Dropdowns
End Property

Private Property Get IMasterSetupPreparation_Variables() As IMasterSetupVariables
    Set IMasterSetupPreparation_Variables = Variables
End Property

Private Property Get IMasterSetupPreparation_HostWorkbook() As Workbook
    Set IMasterSetupPreparation_HostWorkbook = HostWorkbook
End Property
