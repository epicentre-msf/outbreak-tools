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
Private Const VARIABLE_COLUMN_NAME As String = "__Col__Variables"


Private Const DEFAULT_ANCHOR_SHEET As String = "Variables"
Private Const DEFAULT_TABLE_PREFIX As String = "disTab_"
Private Const DEFAULT_LANGUAGE_LIST As String = "__languages"
Private Const DEFAULT_STATUS_LIST As String = "__var_status"
Private Const DEFAULT_VAR_NAME_LIST As String = "PARAMVARNAME"
Private Const DEFAULT_VAR_CHOICE_LIST As String = "PARAMCHOICESLIST"
Private Const DEFAULT_DISEASE_CODE As String = "DISSHEET"

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

Private Const STATUS_DROPDOWN As String = "__var_status"
Private Const DISEASES_DROPDOWN_PROHIBITED As String = "__prohibited_diseases_list"
Private Const DISEASES_DROPODWN As String  = "__diseases_list"
Private Const CHOICES_DROPDOWN As String = "__lst_choices"
Private Const LANGUAGES_DROPDOWN As String = "__data_languages"


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
    hostSheet As Worksheet
    hostIndex As Long
    hostTable As ICustomTable
    hostVariables As IVariables
End Type

Private Type TCurrentHeaders
    varOrder As String
    varSection As String
    varName As String
    varLabel As String
    varChoice As String
    choiceVal As String
    varStatus As String
End Type


Private this As TBuilderState
Private hostHeaders As TCurrentHeaders

'@section Factory
'===============================================================================

'@sub-title Instantiate a sheet builder with provided dependencies.
Public Function Create(ByVal hostWorkbook As Workbook, _
                       ByVal dropdowns As IDropdownLists, _
                       ByVal translator As ITranslationObject, _ 
                       ByVal varObj As IVariables) As IDiseaseSheet

    With New DiseaseSheet
        .Initialise hostWorkbook, dropdowns, translator, varObj
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
                      ByVal translator As ITranslationObject, _ 
                      ByVal varObj As IVariables)

    ValidateInitialise hostWorkbook, dropdowns

    Set this.hostwb = hostWorkbook
    Set this.dropdowns = dropdowns
    Set this.trads = translator
    set this.hostVariables = varObj
    this.initialised = True
End Sub

'@section Worksheet Construction
'===============================================================================

Private Function Build(ByVal diseaseName As String, Optional ByVal languageTag As String) As Worksheet

    ValidateBuildInputs diseaseName, languageTag

    'create the disease worksheet
    CreateDiseaseWorksheet diseaseName

    'Add the table, validations and metadata
    WriteData languageTag
    
    'Format the disease worksheet, add lockings
    FormatDiseaseSheet

    Set Build = this.hostSheet
End Function

'@section Maintenance
'===============================================================================

Public Sub Reset()
    Set this.hostwb = Nothing
    Set this.dropdowns = Nothing
    Set this.trads = Nothing
    this.initialised = False
    Set this.hostSheet = Nothing
    Set this.hostTable = Nothing
    Set DEFAULTDISEASECOLUMNS = Nothing
End Sub

'@section Internal Helpers
'===============================================================================
Private Sub EnsureColumnsInitialized()

    If DEFAULTDISEASECOLUMNS Is Nothing Then

        Set DEFAULTDISEASECOLUMNS = New BetterArray
        
        DEFAULTDISEASECOLUMNS.LowerBound = 1

        hostHeaders.varOrder    = Translate(HEADER_ORDER_KEY, "Variable Order")
        hostHeaders.varName     = Translate(HEADER_NAME_KEY, "Variable Name")
        hostHeaders.varSection  = Translate(HEADER_SECTION_KEY, "Variable Section")
        hostHeaders.varLabel    = Translate(HEADER_LABEL_KEY, "Main Label")
        hostHeaders.varChoice   = Translate(HEADER_CHOICE_KEY, "Choice")
        hostHeaders.choiceVal   = Translate(HEADER_CHOICES_VALUE_KEY, "Choice Values")
        hostHeaders.varStatus   = Translate(HEADER_STATUS_KEY, "Status")


        DEFAULTDISEASECOLUMNS.Push  hostHeaders.varOrder, _
                                    hostHeaders.varName, _
                                    hostHeaders.varSection, _
                                    hostHeaders.varLabel, _
                                    hostHeaders.varChoice, _
                                    hostHeaders.choiceVal, _
                                    hostHeaders.varStatus
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

Private Sub ValidateBuildInputs(ByVal diseaseName As String, Optional ByVal languageTag As String)
    Dim langDrop As BetterArray

    If Not this.initialised Then
        ThrowError ProjectError.ObjectNotInitialized, "Builder has not been initialised"
    End If

    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    If (LenB(languageTag) > 0) Then
        Set langDrop = this.dropdowns.Values(LANGUAGES_DROPDOWN)
        
        If Not langDrop.Includes(languageTag) Then
            ThrowError ProjectError.InvalidArgument, "Language name " & languageTag & "is not present in the list of languages"
        End If

    End If

    If (Len(diseaseName) > 30) Then
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
    Set temp = this.dropdowns.Values(DISEASES_DROPODWN)

    If temp Is Nothing Then
        Set temp = New BetterArray
        temp.LowerBound = 1
    End If

    temp.Push hostsheet.Name

    this.dropdowns.Update temp, DISEASES_DROPODWN
    this.hostIndex = temp.Length
    Set this.hostSheet = hostsheet
End Sub


Private Sub WriteValues()
    
    Dim anchorRng As Range
    Dim mainLabIndex As Long
    Dim choiValIndex As Long
    Dim varNameIndex As Long
    Dim varNameAddress As String
    Dim langAddress As String


    If this.hostSheet Is Nothing Then Exit Sub
    
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN)
    langAddress = anchorRng.Address(True, True)

    anchorRng.Value = Translate(HEADER_LANGUAGE_PROMPT_KEY, "<Select language>")
    
    EnsureColumnsInitialized
    Set anchorRng = anchorRng.Offset(2)
    DEFAULTDISEASECOLUMNS.ToExcelRange anchorRng, TransposeValues:=True

      'Line on values
    Set anchorRng = anchorRng.Offset(1)

    mainLabIndex = ColumnIndexFor(HEADER_LABEL_KEY, "Main Label")
    choiValIndex = ColumnIndexFor(HEADER_CHOICES_VALUE_KEY, "Choice Values")
    varNameIndex = ColumnIndexFor(HEADER_NAME_KEY, "Variable Name")
    If varNameIndex <= 0 Then Exit Sub
    varNameAddress = anchorRng.Cells(1, varNameIndex).Address(False, True)
  
    'Add formulas on specific columns
    If mainLabIndex > 0 Then
        anchorRng.Cells(1, mainLabIndex).Formula = "=MainLabelValue(" & varNameAddress & ", " & langAddress & ")"
    End If

    If choiValIndex > 0 Then
        anchorRng.Cells(1, choiValIndex).Formula = "=ChoiceValues(" & varNameAddress &  ", " & langAddress & " )"
    End If

End Sub

Private Sub CreateDiseaseTable()

    Dim anchorRng As Range 
    Dim tableRange As Range
    Dim dataEnd As Range
    Dim table As ListObject

    If this.hostSheet Is Nothing Then Exit Sub
    
    EnsureColumnsInitialized
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN).Offset(2)
    Set dataEnd = anchorRng.Offset(1, DEFAULTDISEASECOLUMNS.Length - 1)
    Set tableRange = this.hostSheet.Range(anchorRng, dataEnd)
    Set table = this.hostSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    table.Name = DEFAULT_TABLE_PREFIX & Format$(this.hostIndex, "000")
    Set this.hostTable = CustomTable.Create(table)

End Sub


'Add validations on the disease worksheet
Private Sub AddValidations()
    
    Dim anchorRng As Range

    If (this.hostTable Is Nothing) Then Exit Sub
    If (this.hostVariables Is Nothing) Then Exit Sub

    EnsureColumnsInitialized

    'set validation
    this.hostVariables.SetVariableValidation this.hostTable.DataRange(hostHeaders.varName)
    this.hostTable.SetValidation colName:=hostHeaders.varStatus, drop:=this.dropdowns, dropName:=STATUS_DROPDOWN
    this.hostTable.SetValidation colName:=hostHeaders.varChoice, drop:=this.dropdowns, dropName:=CHOICES_DROPDOWN
    
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN)

    this.dropdowns.SetValidation anchorRng, LANGUAGES_DROPDOWN
End Sub


Private Sub WriteData(Optional ByVal languageTag As String)

    WriteValues
    CreateDiseaseTable
    AddValidations
    WriteMetadata languageTag

End Sub

Private Sub WriteMetadata(Optional ByVal languageTag As String)

    Dim storeSheet As IHiddenNames
    Dim storeBook As IHiddenNames
    Dim diseaseCodeName As String
    Dim langVal As String

    If this.hostSheet Is Nothing Then Exit Sub
    langVal = languageTag

    Set storeSheet = HiddenNames.Create(this.hostSheet)
    Set storeBook = HiddenNames.Create(this.hostSheet.Parent)

    'Write the sheet tag
    storeSheet.EnsureName NAME_SHEET_TAG, "disease", HiddenNameTypeString

    'Write the disease index
    storeSheet.EnsureName NAME_INDEX, this.hostIndex , HiddenNameTypeLong

    If (LenB(languageTag) = 0) Then
        Dim langDrop As BetterArray
        Set langDrop = this.dropdowns.Values(LANGUAGES_DROPDOWN)
        langVal = langDrop.Item(langDrop.LowerBound)
    End If

    'Write the language of the disease
    storeSheet.EnsureName NAME_DISLANG, langVal, HiddenNameTypeString

    'Add the disease name to the worksheet
    storeSheet.EnsureName NAME_DISNAME, this.hostSheet.Name, HiddenNameTypeString

    'Add the disease code name
    diseaseCodeName = DEFAULT_DISEASE_CODE & Format$(this.hostIndex, "000")

    'Add the disease code name to the workbook
    storeBook.EnsureName diseaseCodeName, this.hostSheet.Name, HiddenNameTypeString
End Sub


Private Sub FormatDiseaseSheet()
    RemoveGridlines this.hostSheet
    With this.hostSheet
        .Cells.Font.Name = "Calibri"
        .Cells(ANCHORROW, ANCHORCOLUMN).Interior.Color = RGB()
        .Cells(ANCHORROW, ANCHORCOLUMN).Font.Color = vbWhite
    End With

    'Format the table header Range. and interiror color
    With this.hostTable.HeaderRange
        .Font.Bold = True
        .Interior.Color = RGB(237, 242, 248)
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter
    End With

    'Lock required ranges
    this.hostTable.DataRange(hostHeaders.choiceVal).Locked = True
    this.hostTable.DataRange(hostHeaders.varLabel).Locked = True

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

Private Function ColumnIndexFor(ByVal headerKey As String, ByVal fallback As String) As Long
    Dim translated As String

    EnsureColumnsInitialized
    translated = Translate(headerKey, fallback)
    ColumnIndexFor = DEFAULTDISEASECOLUMNS.IndexOf(translated)
    If ColumnIndexFor <= 0 Then
        ColumnIndexFor = DEFAULTDISEASECOLUMNS.IndexOf(fallback)
    End If
End Function

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    If wb Is Nothing Then Exit Function

    For Each sh In wb.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sh
            Exit Function
        End If
    Next sh
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

    If this.trads Is Nothing Then
        retVal = fallback
    Else
        retVal = this.trads.TranslatedValue(key)
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
