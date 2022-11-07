VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Build")
'@ModuleDescription("Object containing specifications for the linelist")
'@IgnoreModule
Implements ILinelistSpecs

Option Explicit

Public Enum TemporarySheetTypes
    TemporaySheetAnalysis = 1
    TemporarySheetDebug = 2
    TemporarySheetMetadata = 3
    TemporarySheetReport = 4
    TemporarySheetList = 5
    TemparySheetImport = 6
    TemporarySingle = 7
End Enum

Private Type TLinelistSpecs
    specswkb As Workbook
End Type

Private this As TLinelistSpecs
Private Const CLASSNAME As String = "LinelistSpecs"
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const CHOICESSHEET As String = "Choices"
Private Const GEOSHEET As String = "Geo"
Private Const PASSWORDSHEET As String = "Password"
Private Const FORMULASHEET As String = "ControleFormule"
Private Const DICTTRANSLATIONSHEET As String = "Translations" 'Dictionary translation sheet
Private Const HLISTTRANSLATIONSHEET As String = "LinelistTranslation" 'linelist translation (forms, shapes, etc.)
Private Const ANALYSISSPECSSHEET As String = "Analysis"
Private Const EXPORTSHEET As String = "Exports"
Private Const FORMATSHEET  As String = "LinelistStyle"

Private Const FORMSLANGUAGECODERANGE As String = "RNG_LLLanguageCode"
Private Const DICTLANGUAGECODERANGE As String = "RNG_DictionaryLanguage"
Private Const designType As String = "DESIGNTYPE"


'Tables
Private Const TABADM1                  As String = "T_ADM1" 'ADM1 Table name
Private Const TABADM2                  As String = "T_ADM2" 'ADM2 Table name
Private Const TABADM3                  As String = "T_ADM3" 'ADM3 Table name
Private Const TABADM4                  As String = "T_ADM4" 'ADM4 Table name
Private Const TABHF                    As String = "T_HF" 'Health Facility Table
Private Const TABNAMES                 As String = "T_NAMES"
Private Const TABHISTOGEO              As String = "T_HistoGeo" 'Historic data for the geo
Private Const TABHISTOHF               As String = "T_HistoHF" 'Historic data for the Health Facility
Private Const TABGEOMETADATA           As String = "T_Metadata"

'Tables for traductions

Private Const TABTRADLLMSG             As String = "T_TradLLMsg"
Private Const TABTRADLLSHAPES          As String = "T_TradLLShapes"
Private Const TABTRADLLFORMS           As String = "T_TradLLForms"
Private Const TABLLLANG                As String = "T_LLLang"
Private Const TABTRANSLATION           As String = "Tab_Translations"

'Formulas and functions tables
Private Const TABEXCELFUNCTIONS        As String = "T_XlsFonctions" 'Excel functions to keep in formulas
Private Const TABASCII                 As String = "T_ascii" 'Ascii characters table

Public Function Create(ByVal specswkb As Workbook) As ILinelistSpecs
    Dim counter As Long
    Dim sheetName As String

    If specswkb Is Nothing Then ThrowErrorUnableInstanciate "specification workbook does not exists"
    
    'Check worksheets requirements
    CheckSheetsRequirements specswkb
    'Check requirements for Listobject tables
    CheckTablesRequirements specswkb

    'Check Ranges requirements
    CheckRangesRequirements specswkb

    'At this step everything needed of creating the Linelist exists
    With New LinelistSpecs
        Set .Wkb = specswkb
        Set Create = .Self
    End With
End Function

'Initialize the array of sheets
Private Sub CheckSheetsRequirements(ByVal specswkb As Workbook)
    Dim counter As Long
    Dim sheetName As String

    Dim sheetsArray As BetterArray
    Set sheetsArray = New BetterArray
    sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                     PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                     EXPORTSHEET, FORMATSHEET

    For counter = sheetsArray.LowerBound To sheetsArray.UpperBound
        sheetName = sheetsArray.Item(counter)
        If Not sheetExists(specswkb, sheetName) Then ThrowErrorUnableInstanciate " Worksheet " & sheetName & " Not found!"
    Next
End Sub

Private Sub CheckTablesRequirements(ByVal specswkb As Workbook)
    Dim counter As Long
    Dim sh As Worksheet
    Dim loArray As BetterArray
    Set sh = specswkb.Worksheets(HLISTTRANSLATIONSHEET)
    Set loArray = New BetterArray

    'Forms translation
    loArray.Push TABTRADLLMSG, TABTRADLLSHAPES, TABTRADLLFORMS, TABLLLANG
    LoExists sh, loArray

    'Dictionary translation
    loArray.Clear
    loArray.Push TABTRANSLATION
    Set sh = specswkb.Worksheets(DICTTRANSLATIONSHEET)
    LoExists sh, loArray
    
    'Geobase
    loArray.Clear
    Set sh = specswkb.Worksheets(GEOSHEET)
    loArray.Push TABADM1, TABADM2, TABADM3, TABADM4, TABHF, TABNAMES, _
                 TABHISTOGEO, TABHISTOHF, TABGEOMETADATA

    LoExists sh, loArray

    'Formulas
    loArray.Clear
    Set sh = specswkb.Worksheets(FORMULASHEET)
    loArray.Push TABEXCELFUNCTIONS, TABASCII
    LoExists sh, loArray
End Sub

Private Sub CheckRangesRequirements(ByVal specswkb As Workbook)
    Dim sh As Worksheet

    'Ranges in the LinelistTranslation sheet
    Set sh = specswkb.Worksheets(HLISTTRANSLATIONSHEET)
    RangeExists sh, DICTLANGUAGECODERANGE
    RangeExists sh, FORMSLANGUAGECODERANGE
    
    Set sh = specswkb.Worksheets(FORMATSHEET)
    RangeExists sh, designType
End Sub

'Check if a sheet Exists
Private Function sheetExists(ByVal specswkb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet                          'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set sh = specswkb.Worksheets(sheetName)
    sheetExists = (Not sh Is Nothing)
    On Error GoTo 0
End Function

Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    If rng Is Nothing Then
        ThrowErrorUnableInstanciate "Range " & rngName & " does not exists in worksheet " & sh.Name
    End If
    On Error GoTo 0
End Sub

Private Sub LoExists(ByVal sh As Worksheet, ByVal loNames As BetterArray)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String
    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(loNames)
        If Lo Is Nothing Then
            ThrowErrorUnableInstanciate "Table " & lName & " does not exists in worksheet" & sh.Name
        End If
        On Error GoTo 0
    Next
End Sub

Public Property Get Wkb() As Workbook
    Set Wkb = this.specswkb
End Property

Public Property Set Wkb(ByVal specswkb As Workbook)
    Set this.specswkb = specswkb
End Property

Public Property Get Self() As ILinelistSpecs
    Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
    'Create the dictionary
    Dim dict As ILLdictionary
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(DICTIONARYSHEET)

    'StartRow is 1, start column is 1
    Set dict = LLdictionary.Create(sh, 1, 1)
    Set Dictionary = dict
End Property

Public Property Get Choices() As ILLchoice
    Dim choi As ILLchoice
    Dim sh As Worksheet

    Set sh = Wkb.Worksheets(CHOICESSHEET)

    Set choi = LLchoice.Create(sh, 1, 1)
    Set Choices = choi
End Property

Private Property Get FormLang() As String
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(HLISTTRANSLATIONSHEET)
    FormLang = sh.Range(FORMSLANGUAGECODERANGE).Value
End Property

Private Property Get DictLang() As String
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(HLISTTRANSLATIONSHEET)
    DictLang = sh.Range(DICTLANGUAGECODERANGE).Value
End Property

'Translation Object for Dictionary, choices, analysis Translations
Public Property Get DictTransObject() As ITranslation
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim lang As String
    
    Set sh = Wkb.Worksheets(DICTTRANSLATIONSHEET)
    Set Lo = sh.ListObjects(1)
    lang = DictLang()

    Set DictTransObject = Translation.Create(Lo, lang)
End Property

'Test if the translation form exists
Public Property Get FormTransObject(Optional ByVal loName As String = "T_TradLLMsg") As ITranslation

    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim lang As String

    Set sh = Wkb.Worksheets(HLISTTRANSLATIONSHEET)
    Set Lo = sh.ListObjects(loName)
    lang = FormLang()

    Set FormTransObject = Translation.Create(Lo, lang)
End Property

Public Property Get FormDataObject() As IFormulaData
    Dim sh As Worksheet
    Dim excelLo As ListObject                    'Excel Functions listObject
    Dim charLo As ListObject                     'special characters listObject

    Set sh = Wkb.Worksheets(FORMULASHEET)
    Set excelLo = sh.ListObjects(TABEXCELFUNCTIONS)
    Set charLo = sh.ListObjects(TABASCII)

    Set FormDataObject = FormulaData.Create(sh, excelLo, charLo)
End Property

Public Property Get DesignFormat() As ILLFormat
    Dim sh As Worksheet
    Dim designName As String

    Set sh = Wkb.Worksheets(FORMATSHEET)
    designName = sh.Range(designType).Value

    'Will take in account the other aspecs
    Set DesignFormat = LLFormat.Create(sh, designName)
End Property

Public Property Get TemporarySheetName(ByVal scope As TemporarySheetTypes) As String

    Select Case scope
    Case TemporaySheetAnalysis
        TemporarySheetName = "ana_temp__"
    Case TemporarySheetDebug
        TemporarySheetName = "debug__"
    Case TemporarySheetMetadata
        TemporarySheetName = "metadata__"
    Case TemporarySheetReport
        TemporarySheetName = "import_temp__"
    Case TemporarySheetList
        TemporarySheetName = "dropdown_lists__"
    Case TemparySheetImport
        TemporarySheetName = "import_temp__"
    Case TemporarySingle
        TemporarySheetName = "temp__"
    End Select
End Property

'Get Categories for one variable
Private Function Categories(ByVal varName As String) As BetterArray

    Dim categoryData As BetterArray
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim categoryName As String
    Dim categoryType As String

    Set dict = Dictionary()
    Set choi = Choices()
    Set categoryData = New BetterArray
    Set vars = LLVariables.Create(dict)
    categoryType = vars.Value(varName:=varName, colName:="control")
    If categoryType = "case_when" Then
        categoryName = "__case_when_" & varName
    ElseIf categoryType = "choices" Then
        categoryName = vars.Value(varName:=varName, colName:="control details")
    End If
    'Test the case_when
    Set categoryData = choi.Categories(categoryName)
    Set Categories = categoryData.Clone()

End Function

Private Sub Prepare()

    Const CATEGORYPREFIX As String = "__case_when_"
    Dim PreservedSheetNames As BetterArray
    
    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim vars As ILLVariables
    Dim trans As ITranslation
    Dim dicttrans As ITranslation
    Dim counter As Long
    
    Dim casewhenVars As BetterArray
    Dim casewhenObject As ICaseWhen
    Dim casewhenVariableName As String
    Dim casewhenFormula As String
    
    
    Set dict = Dictionary()
    Set choi = Choices()
    Set trans = FormTransObject()
    Set dicttrans = DictTransObject()
    
    Set PreservedSheetNames = New BetterArray
    PreservedSheetNames.LowerBound = 1


    'Preserved names for worksheet
    PreservedSheetNames.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                             PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                             EXPORTSHEET, "temp_", "list_auto_", "debug___", "metadata", "ana_temp_", "import_temp_", _
                             trans.TranslatedValue("LLSHEET_Admin"), trans.TranslatedValue("LLSHEET_Analysis"), _
                             trans.TranslatedValue("LLSHEET_SpatialAnalysis"), trans.TranslatedValue("LLSHEET_TemporalAnalysis")
    
    'Translate the dictionary
    dict.Translate dicttrans
    
    'Prepare the dictionary
    dict.Prepare PreservedSheetNames, tablePrefix:="table"
    'Need to prepare the dictionary before loading the variable class
    Set vars = LLVariables.Create(dict)
 
    'Add case_when to choices
    Set casewhenVars = dict.data.FilterData(varName:="control", criteriaName:="case_when", returnedColumnName:="variable name")
    For counter = casewhenVars.LowerBound To casewhenVars.UpperBound
        casewhenVariableName = casewhenVars.Item(counter)
        casewhenFormula = vars.Value(varName:=casewhenVariableName, colName:="control details")

        Set casewhenObject = CaseWhen.Create(casewhenFormula)
        'Test validity before parsing categories to choices
        If casewhenObject.valid() Then
            choi.AddChoice choiceName:=CATEGORYPREFIX & casewhenVariableName, CategoriesData:=casewhenObject.Categories
        End If
    Next

    'Translate the choices
    choi.Translate dicttrans
    
    'Sort the exports
    choi.Sort

End Sub

'Errors
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Implement the interfaces

Public Property Get ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = Dictionary()
End Property

Public Property Get ILinelistSpecs_Choices() As ILLchoice
    Set ILinelistSpecs_Choices = Choices()
End Property

'Categories of one variable
Public Function ILinelistSpecs_Categories(ByVal varName As String) As BetterArray
    Set ILinelistSpecs_Categories = Categories(varName)
End Function

Public Sub ILinelistSpecs_Prepare()
    Prepare
End Sub

Public Property Get ILinelistSpecs_TransObject(Optional ByVal loName As String = "T_TradLLMsg") As ITranslation
    Set ILinelistSpecs_TransObject = FormTransObject(loName)
End Property

'Formula data object
Public Property Get ILinelistSpecs_FormDataObject() As IFormulaData
    Set ILinelistSpecs_FormDataObject = FormDataObject()
End Property

Public Property Get ILinelistSpecs_TemporarysheetName(ByVal scope As TemporarySheetTypes) As String
    ILinelistSpecs_TemporarysheetName = TemporarySheetName(scope)
End Property

Public Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = DesignFormat()
End Property


