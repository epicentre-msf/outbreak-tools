VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds disease worksheets with translated headers, validations, and formatting"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Builds disease worksheets with translated headers, validations, and formatting")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseSheet

'@section Constants
'===============================================================================

Private Const NAME_SHEET_TAG As String = "sheetTag"
Private Const NAME_DISNAME As String = "__Var_DISNAME"
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_INDEX As String = "__Var_DISINDEX"

Private Const DEFAULT_ANCHOR_SHEET As String = "Variables"
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


Private Const DISEASES_DROPDOWN_PROHIBITED As String = "__prohibited_diseases_list"

Private DEFAULTDISEASECOLUMNS As BetterArray

Private Const ANCHORROW As Long = 2
Private Const ANCHORCOLUMN As Long =2

'@section State
'===============================================================================

Private Type TBuilderState
    hostwb As Workbook
    dropdowns As IDropdownLists
    trads As ITranslationObject
    initialised As Boolean
    activeHostsheet As Worksheet
    activeIndex As Long
    activeTabwrapper As ICustomTable
End Type


Private this As TBuilderState

'@section Factory
'===============================================================================

'@sub-title Instantiate a sheet builder with provided dependencies.
Public Function Create(ByVal hostWorkbook As Workbook, _
                       ByVal dropdowns As IDropdownLists, _
                       ByVal translator As ITranslationObject) As IDiseaseSheet

    With New DiseaseSheet
        .Initialise hostWorkbook, dropdowns, translator
        Set Create = .Self
    End With
End Function

'@sub-title Return the current instance as its interface.
Public Property Get Self() As IDiseaseSheet
    Set Self = Me
End Property

'@section Configuration
'===============================================================================

Public Sub Initialise(ByVal hostWorkbook As Workbook, _
                      ByVal dropdowns As IDropdownLists, _
                      ByVal translator As ITranslationObject)

    ValidateInitialise hostWorkbook, dropdowns

    Set this.hostwb = hostWorkbook
    Set this.dropdowns = dropdowns
    Set this.trads = translator
    this.initialised = True
End Sub

'@section Worksheet Construction
'===============================================================================

Private Function Build(ByVal diseaseName As String) As Worksheet

    Dim config As TBuilderOptions

    ValidateBuildInputs diseaseName
    CreateDiseaseWorksheet diseaseName
    WriteMetadata
    AddInternalStore
    ConfigureValidations diseaseSheet, config, messageSelectValue
    CreateDiseaseTable diseaseSheet, config, diseaseIndex
    FormatDiseaseSheet diseaseSheet

    Set Build = diseaseSheet
End Function

'@section Maintenance
'===============================================================================

Public Sub Reset()
    Set this.hostwb = Nothing
    Set this.dropdowns = Nothing
    Set this.trads = Nothing
    this.initalised = False
    Set this.activeHostsheet = Nothing
End Sub

'@section Internal Helpers
'===============================================================================
Private Sub EnsureColumnsInitialized()

    If DEFAULTDISEASECOLUMNS Is Nothing Then

        Set DEFAULTDISEASECOLUMNS = New BetterArray
        
        DEFAULTDISEASECOLUMNS.LowerBound = 1
        DEFAULTDISEASECOLUMNS.Push   Translate(HEADER_ORDER_KEY, "Variable Order"), _
                    Translate(HEADER_NAME_KEY, "Variable Name"), _
                    Translate(HEADER_SECTION_KEY, "Variable Section"), _
                    Translate(HEADER_LABEL_KEY, "Main Label"), _
                    Translate(HEADER_CHOICE_KEY, "Choice"), _
                    Translate(HEADER_CHOICES_VALUE_KEY, "Choice Values"), _
                    Translate(HEADER_STATUS_KEY, "Status")
    End If
End Sub



Private Sub ValidateInitialise(ByVal hostWorkbook As Workbook, ByVal dropdowns As IDropdownLists)

    If hostWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If

    If dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dropdown lists cannot be Nothing"
    End If
End Sub

Private Sub ValidateBuildInputs(ByVal diseaseName As String)
    If Not this.initialised Then
        ThrowError ProjectError.ObjectNotInitialized, "Builder has not been initialised"
    End If


    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    If (LenB(diseaseName) > 30) Then
        ThrowError ProjectError.ObjectNotInitialized, "Disease name should be less than 30 characters long (including spaces)"
    End If

    'Testing if the name of the disease is correct
    Dim temp As BetterArray
    Set temp = this.dropdowns.Values(DISEASES_DROPDOWN_PROHIBITED)

    If temp.Includes(diseaseName) Then
        ThrowError ProjectError.InvalidArgument, _ 
         "The name " & diseaseName & " is incorrect because it is reserved. Please use another name"
    End If

    Set temp = this.dropdowns.Values(DISEASES_DROPODWN)
    
    If temp.Includes(diseaseName) Then
        ThrowError ProjectError.InvalidArgument, _ 
         "The name " & diseaseName & " already exists, please choose another name."
    End If
End Sub


Private Sub CreateDiseaseWorksheet(ByVal diseaseName As String)
    Dim anchor As Worksheet
    Dim hostwb As Workbook
    Dim hostsheet As Worksheet
 
    Set hostwb = this.hostwb

    On Error Resume Next
    Set anchor = hostwb.Worksheets(DEFAULT_ANCHOR_SHEET)
    On Error GoTo 0

    If (anchor Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, _ 
         "Anchor worksheet " & DEFAULT_ANCHOR_SHEET & "not found"
    End If

    Set hostsheet = hostwb.Worksheets.Add(Before:=anchor)
    hostsheet.Name = diseaseName
    UpdateInternalElements hostsheet
End Sub

Private Sub UpdateInternalElements(ByVal hostsheet As Worksheet)
    
    Dim temp As BetterArray
    Dim store As IHiddenNames
    Dim hostSh As Worksheet
    Dim hostwb As Workbook

    Set store = HiddenNames.Create(hostsheet)
    Set temp = this.dropdowns.Values(DISEASES_DROPODWN)
    temp.Push hosthseet.Name

    this.dropdowns.Update temp, DISEASES_DROPODWN
    this.activeIndex = temp.Length

    Set this.activeHostSheet = hostsheet
    'stock the index
    store.EnsureName NAME_INDEX, this.activeIndex, HiddenNameTypeLong

    'stock the language
    store.EnsureName NAME_DISLANG, this.trads.Language, HiddenNameTypeString
End Sub



Private Sub WriteValues()
    
    Dim anchorRng As Range
    Dim mainLabIndex As Long
    Dim choiValIndex As Long
    Dim varNameIndex As Long
    Dim varNameAddress As String


    If this.activeHostSheet Is Nothing Then Exit Sub
    
    Set anchorRng = this.activeHostSheet.Cells(ANCHORROW, ANCHORCOLUMN)

    anchorRng.Value = Translate(HEADER_LANGUAGE_PROMPT_KEY, "<Select language>")
    
    EnsureColumnsInitialized
    Set anchorRng = anchorRng.Offset(2)
    DEFAULTDISEASECOLUMNS.ToExcelRange anchorRng, TransposeValues:=True

      'Line on values
    Set anchorRng = anchorRng.Offset(1)

    mainLabIndex = DEFAULTDISEASECOLUMNS.IndexOf(Translate(HEADER_LABEL_KEY, "Main Label"))
    choiValIndex = DEFAULTDISEASECOLUMNS.IndexOf(Translate(HEADER_CHOICES_VALUE_KEY, "Choice Values"))
    varNameIndex = EFAULTDISEASECOLUMNS.IndexOf(Translate(HEADER_NAME_KEY, "Variable Name"))
    varNameAddress = anchorRng.Cells(1, varNameIndex).Address(False, True)
  

    'Add formulas on specific columns
    anchorRng.Cells(1, mainLabIndex).Formula = "=MainLabelValue(" & varNameAddress & ")"
    anchorRng.Cells(1, choiValIndex).Formula = "=ChoiceValues(" & varNameAddress & ")"
End Sub



Private Sub CreateDiseaseTable()

    Dim anchorRng As Range 
    Dim tableRange As Range
    Dim headerRow As Long
    Dim lastRow As Long
    Dim table As ListObject
    Dim tableName As String

    If this.activeHostSheet Is Nothing Then Exit Sub
    
    EnsureColumnsInitialized
    Set anchorRng = this.activeHostSheet.Cells(ANCHORROW, ANCHORCOLUMN).Offset(2)
    Set tableRange = sheet.Range(anchorRng, anchorRng.Offset(1, (DEFAULTDISEASECOLUMNS.Length - 1)))
    Set table = sheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = DEFAULT_TABLE_PREFIX & Format$(this.activeIndex, "000")
    Set this.activeTabwrapper = CustomTable.Create(table)
End Sub


'Add validations on the disease worksheet
Private Sub AddValidations()

    
    'Add validation on variable Name


    'Add validation on status


End Sub


Private Sub WriteMetadata(ByVal diseasesheet As Worksheet, _
                          config As TBuilderOptions, _
                          ByVal languageValue As String, _
                          ByVal selectMessage As String)

    WriteValues
    CreateDiseaseTable
    AddValidations

    Set hostwb = this.hostwb
    
    sheet.Cells(5, 4).Validation.Delete
    ConfigureValidation sheet.Cells(5, 4), config.varChoiceListName, selectMessage, False
End Sub

Private Sub AddInternalStore(ByVal hostsheet As Worksheet, config As TBuilderOptions)

    Dim storesheet As IHiddenNames
    Dim storebook As IHiddenNames

    Set storesheet = HiddenNames.Create(hostsheet)
    Set storesbook = HiddenNames.Create(hostsheet.Parent)

    'Add the disease index
    

        sheet.Cells(2, 3).Value = diseaseIndex
    sheet.Cells(2, 4).Value = config.diseaseCode



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

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    For Each sh In workbook.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sh
            Exit Function
        End If
    Next
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
    Dim retVal As String

    If( this.translator Is Nothing) Then
        retVal = fallback
    Else
        retVal = this.translator.TranslatedValue(key)
        retVal = IIf(retVal = key, fallback, retVal)
    End If

    Translate = retVal
End Function

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseSheet", errMessage
End Sub

'@section Interface
'===============================================================================

Private Sub IDiseaseSheet_Initialise(ByVal hostWorkbook As Workbook, ByVal dropdowns As IDropdownLists, ByVal translations As ITranslationObject)
    Initialise hostWorkbook, dropdowns, translations
End Sub

Private Function IDiseaseSheet_Build(ByVal diseaseName As String, ByVal diseaseIndex As Long, Optional ByVal languageTag As String = vbNullString, Optional ByVal options As BetterArray) As Worksheet
    Set IDiseaseSheet_Build = Build(diseaseName, diseaseIndex, languageTag, options)
End Function

Private Sub IDiseaseSheet_Reset()
    Reset
End Sub
