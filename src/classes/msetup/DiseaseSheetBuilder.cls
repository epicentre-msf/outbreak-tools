VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseSheetBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Builds disease worksheets with translated headers, validations, and formatting"

'@Folder("Msetup")
'@ModuleDescription("Builds disease worksheets with translated headers, validations, and formatting")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseSheetBuilder

'@section Constants
'===============================================================================

Private Const DEFAULT_ANCHOR_SHEET As String = "Variables"
Private Const DEFAULT_LANGUAGE_PREFIX As String = "disLang_"
Private Const DEFAULT_TABLE_PREFIX As String = "disTab_"
Private Const DEFAULT_LANGUAGE_LIST As String = "__languages"
Private Const DEFAULT_STATUS_LIST As String = "__var_status"
Private Const DEFAULT_VAR_NAME_LIST As String = "PARAMVARNAME"
Private Const DEFAULT_VAR_CHOICE_LIST As String = "PARAMCHOICESLIST"
Private Const DEFAULT_DISEASE_CODE As String = "DISSHEET"
Private Const DEFAULT_TABLE_ROWS As Long = 15

Private Const HEADER_ORDER_KEY As String = "varOrder"
Private Const HEADER_SECTION_KEY As String = "varSection"
Private Const HEADER_NAME_KEY As String = "varName"
Private Const HEADER_LABEL_KEY As String = "varLabel"
Private Const HEADER_CHOICE_KEY As String = "varChoice"
Private Const HEADER_CHOICES_VALUE_KEY As String = "choiceVal"
Private Const HEADER_STATUS_KEY As String = "varStatus"
Private Const HEADER_LANGUAGE_PROMPT_KEY As String = "infoSelectLang"
Private Const MESSAGE_SELECT_VALUE_KEY As String = "selectValue"
Private Const MESSAGE_LANGUAGE_ERROR_KEY As String = "errLang"

'@section State
'===============================================================================

Private Type TBuilderState
    workbook As Workbook
    dropdowns As IDropdownLists
    translations As ITranslationObject
    cache As IDiseaseTranslationCache
End Type

Private Type TBuilderOptions
    anchorSheetName As String
    languagePrefix As String
    tablePrefix As String
    languageListName As String
    statusListName As String
    varNameListName As String
    varChoiceListName As String
    diseaseCode As String
End Type

Private this As TBuilderState

'@section Factory
'===============================================================================

'@sub-title Instantiate a sheet builder with provided dependencies.
Public Function Create(ByVal hostWorkbook As Workbook, _
                       ByVal dropdowns As IDropdownLists, _
                       ByVal translations As ITranslationObject) As IDiseaseSheetBuilder

    With New DiseaseSheetBuilder
        .Initialise hostWorkbook, dropdowns, translations
        Set Create = .Self
    End With
End Function

'@sub-title Return the current instance as its interface.
Public Property Get Self() As IDiseaseSheetBuilder
    Set Self = Me
End Property

'@section Configuration
'===============================================================================

Public Sub Initialise(ByVal hostWorkbook As Workbook, _
                      ByVal dropdowns As IDropdownLists, _
                      ByVal translations As ITranslationObject)

    ValidateInitialise hostWorkbook, dropdowns

    Set this.workbook = hostWorkbook
    Set this.dropdowns = dropdowns
    Set this.translations = translations
    Set this.cache = DiseaseTranslationCache.Create(translations)
End Sub

'@section Worksheet Construction
'===============================================================================

Public Function Build(ByVal diseaseName As String, _
                      ByVal diseaseIndex As Long, _
                      Optional ByVal languageTag As String = vbNullString, _
                      Optional ByVal options As BetterArray) As Worksheet

    Dim config As TBuilderOptions
    Dim diseaseSheet As Worksheet
    Dim messageSelectValue As String
    Dim languagePrompt As String

    ValidateBuildInputs diseaseName, diseaseIndex
    config = ResolveOptions(options)

    languagePrompt = Translate(HEADER_LANGUAGE_PROMPT_KEY, "Select language")
    messageSelectValue = Translate(MESSAGE_SELECT_VALUE_KEY, "Select a value")

    Set diseaseSheet = CreateDiseaseWorksheet(config, diseaseName)
    WriteMetadata diseaseSheet, diseaseIndex, config, IIf(LenB(languageTag) = 0, languagePrompt, languageTag), messageSelectValue
    ConfigureValidations diseaseSheet, config, messageSelectValue
    CreateDiseaseTable diseaseSheet, config, diseaseIndex
    FormatDiseaseSheet diseaseSheet

    Set Build = diseaseSheet
End Function

'@section Maintenance
'===============================================================================

Public Sub Reset()
    Set this.workbook = Nothing
    Set this.dropdowns = Nothing
    Set this.translations = Nothing
    Set this.cache = Nothing
End Sub

'@section Internal Helpers
'===============================================================================

Private Sub ValidateInitialise(ByVal hostWorkbook As Workbook, ByVal dropdowns As IDropdownLists)
    If hostWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If

    If dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dropdown lists cannot be Nothing"
    End If
End Sub

Private Sub ValidateBuildInputs(ByVal diseaseName As String, ByVal diseaseIndex As Long)
    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    If diseaseIndex <= 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease index must be greater than zero"
    End If

    If this.workbook Is Nothing Or this.dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Builder has not been initialised"
    End If
End Sub

Private Function ResolveOptions(ByVal options As BetterArray) As TBuilderOptions
    Dim result As TBuilderOptions

    result.anchorSheetName = DEFAULT_ANCHOR_SHEET
    result.languagePrefix = DEFAULT_LANGUAGE_PREFIX
    result.tablePrefix = DEFAULT_TABLE_PREFIX
    result.languageListName = DEFAULT_LANGUAGE_LIST
    result.statusListName = DEFAULT_STATUS_LIST
    result.varNameListName = DEFAULT_VAR_NAME_LIST
    result.varChoiceListName = DEFAULT_VAR_CHOICE_LIST
    result.diseaseCode = DEFAULT_DISEASE_CODE

    If options Is Nothing Then
        ResolveOptions = result
        Exit Function
    End If

    Dim idx As Long
    Dim key As String

    For idx = options.LowerBound To options.UpperBound Step 2
        key = LCase$(CStr(options.Item(idx)))
        Select Case key
            Case "anchorsheet": result.anchorSheetName = CStr(options.Item(idx + 1))
            Case "languageprefix": result.languagePrefix = CStr(options.Item(idx + 1))
            Case "tableprefix": result.tablePrefix = CStr(options.Item(idx + 1))
            Case "languagelist": result.languageListName = CStr(options.Item(idx + 1))
            Case "statuslist": result.statusListName = CStr(options.Item(idx + 1))
            Case "varnamelist": result.varNameListName = CStr(options.Item(idx + 1))
            Case "varchoicelist": result.varChoiceListName = CStr(options.Item(idx + 1))
            Case "diseasecode": result.diseaseCode = CStr(options.Item(idx + 1))
        End Select
    Next idx

    ResolveOptions = result
End Function

Private Function CreateDiseaseWorksheet(ByVal config As TBuilderOptions, ByVal diseaseName As String) As Worksheet
    Dim anchor As Worksheet
    Dim workbook As Workbook

    Set workbook = this.workbook
    Set anchor = ResolveAnchorSheet(workbook, config.anchorSheetName)

    workbook.Worksheets.Add Before:=anchor
    workbook.ActiveSheet.Name = diseaseName

    Set CreateDiseaseWorksheet = workbook.Worksheets(diseaseName)
End Function

Private Function ResolveAnchorSheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    Set ResolveAnchorSheet = FindWorksheet(workbook, sheetName)
    If ResolveAnchorSheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
            "Anchor worksheet '" & sheetName & "' was not found in the workbook"
    End If
End Function

Private Sub WriteMetadata(ByVal sheet As Worksheet, _
                          ByVal diseaseIndex As Long, _
                          ByVal config As TBuilderOptions, _
                          ByVal languageValue As String, _
                          ByVal selectMessage As String)

    Dim workbook As Workbook

    Set workbook = this.workbook

    sheet.Cells(2, 2).Value = languageValue
    AssignName workbook, sheet.Cells(2, 2), config.languagePrefix & Format$(diseaseIndex, "0")

    sheet.Cells(2, 3).Value = diseaseIndex
    sheet.Cells(2, 4).Value = config.diseaseCode

    sheet.Cells(4, 2).Value = Translate(HEADER_ORDER_KEY, "Variable Order")
    sheet.Cells(4, 3).Value = Translate(HEADER_SECTION_KEY, "Variable Section")
    sheet.Cells(4, 4).Value = Translate(HEADER_NAME_KEY, "Variable Name")
    sheet.Cells(4, 5).Value = Translate(HEADER_LABEL_KEY, "Main Label")
    sheet.Cells(4, 6).Value = Translate(HEADER_CHOICE_KEY, "Control")
    sheet.Cells(4, 7).Value = Translate(HEADER_CHOICES_VALUE_KEY, "Choices")
    sheet.Cells(4, 8).Value = Translate(HEADER_STATUS_KEY, "Status")

    sheet.Cells(5, 4).Validation.Delete
    ConfigureValidation sheet.Cells(5, 4), config.varChoiceListName, selectMessage, False
End Sub

Private Sub ConfigureValidations(ByVal sheet As Worksheet, _
                                 ByVal config As TBuilderOptions, _
                                 ByVal selectMessage As String)

    Dim dropdowns As IDropdownLists

    Set dropdowns = this.dropdowns

    dropdowns.SetValidation sheet.Cells(2, 2), config.languageListName, "error", _
        Translate(MESSAGE_LANGUAGE_ERROR_KEY, selectMessage)

    dropdowns.SetValidation sheet.Cells(5, 6), config.statusListName, "error", selectMessage

    ConfigureValidation sheet.Cells(5, 2), config.varNameListName, selectMessage, True
End Sub

Private Sub ConfigureValidation(ByVal target As Range, _
                                ByVal listName As String, _
                                ByVal errorMessage As String, _
                                ByVal provideInput As Boolean)

    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="= " & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = IIf(provideInput, errorMessage, vbNullString)
        .ErrorMessage = errorMessage
        .ShowInput = provideInput
        .ShowError = True
    End With
End Sub

Private Sub CreateDiseaseTable(ByVal sheet As Worksheet, _
                               ByVal config As TBuilderOptions, _
                               ByVal diseaseIndex As Long)

    Dim tableRange As Range
    Dim headerRow As Long
    Dim lastRow As Long
    Dim table As ListObject
    Dim tableName As String

    headerRow = 4
    lastRow = headerRow + DEFAULT_TABLE_ROWS

    Set tableRange = sheet.Range(sheet.Cells(headerRow, 2), sheet.Cells(lastRow, 8))

    Set table = sheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    tableName = config.tablePrefix & Format$(diseaseIndex, "0")
    table.Name = tableName
End Sub

Private Sub FormatDiseaseSheet(ByVal sheet As Worksheet)
    RemoveGridlines sheet
    With sheet
        .Cells.Font.Name = "Calibri"
        .Cells.Columns.AutoFit
        .Rows(4).Font.Bold = True
        .Rows(4).Interior.Color = RGB(237, 242, 248)
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub RemoveGridlines(ByVal sheet As Worksheet)
    Dim view As WorksheetView
    For Each view In sheet.Parent.Windows(1).SheetViews
        If view.Sheet.Name = sheet.Name Then
            view.DisplayGridlines = False
            Exit For
        End If
    Next view
End Sub

Private Function FindWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In workbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet
End Function

Private Sub EnsureNameDeleted(ByVal workbook As Workbook, ByVal nameValue As String)
    Dim namedRange As Name

    For Each namedRange In workbook.Names
        If StrComp(namedRange.Name, nameValue, vbTextCompare) = 0 Then
            namedRange.Delete
            Exit For
        End If
    Next namedRange
End Sub

Private Sub AssignName(ByVal workbook As Workbook, ByVal target As Range, ByVal nameValue As String)
    EnsureNameDeleted workbook, nameValue
    workbook.Names.Add Name:=nameValue, RefersTo:=target
End Sub

Private Function Translate(ByVal key As String, ByVal fallback As String) As String
    If this.cache Is Nothing Then
        Translate = fallback
    Else
        Translate = this.cache.Resolve(key, fallback)
    End If
End Function

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseSheetBuilder", errMessage
End Sub

'@section Interface
'===============================================================================

Private Sub IDiseaseSheetBuilder_Initialise(ByVal hostWorkbook As Workbook, ByVal dropdowns As IDropdownLists, ByVal translations As ITranslationObject)
    Initialise hostWorkbook, dropdowns, translations
End Sub

Private Function IDiseaseSheetBuilder_Build(ByVal diseaseName As String, ByVal diseaseIndex As Long, Optional ByVal languageTag As String = vbNullString, Optional ByVal options As BetterArray) As Worksheet
    Set IDiseaseSheetBuilder_Build = Build(diseaseName, diseaseIndex, languageTag, options)
End Function

Private Sub IDiseaseSheetBuilder_Reset()
    Reset
End Sub
