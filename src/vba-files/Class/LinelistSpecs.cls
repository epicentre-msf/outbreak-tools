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
    TemporarySheetAnalysis = 1
    TemporarySheetDebug = 2
    TemporarySheetMetadata = 3
    TemporarySheetList = 4
    TemporarySheetImport = 5
    TemporarySingle = 6
End Enum

Private Type TLinelistSpecs
    specswkb As Workbook
End Type

Private this As TLinelistSpecs
Private Const CLASSNAME                 As String = "LinelistSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const CHOICESSHEET              As String = "Choices"
Private Const GEOSHEET                  As String = "Geo"
Private Const PASSWORDSHEET             As String = "Password"
Private Const FORMULASHEET              As String = "ControleFormule"
Private Const DICTTRANSLATIONSHEET      As String = "Translations" 'Dictionary translation sheet
Private Const HLISTTRANSLATIONSHEET     As String = "LinelistTranslation" 'linelist translation (forms, shapes, etc.)
Private Const ANALYSISSPECSSHEET        As String = "Analysis"
Private Const EXPORTSHEET               As String = "Exports"
Private Const FORMATSHEET               As String = "LinelistStyle"

'Name of the range for design name in the design sheet
Private Const designType                As String = "DESIGNTYPE"

Public Function Create(ByVal specswkb As Workbook) As ILinelistSpecs

    Dim sheetName As String
    If specswkb Is Nothing Then ThrowErrorUnableInstanciate "specification workbook does not exists"
    'Check worksheets requirements
    CheckSheetsRequirements specswkb
    'At this step everything needed of creating the Linelist exists
    With New LinelistSpecs
        Set .Wkb = specswkb
        Set Create = .Self
    End With

End Function


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

Public Property Get Trads() As ILLTranslations
    Dim trans As ILLTranslations
    Dim shtransLL As Worksheet
    Dim shtransDict As Worksheet

    Set shtransLL = Wkb.Worksheets(HLISTTRANSLATIONSHEET)
    Set shtransDict = Wkb.Worksheets(DICTTRANSLATIONSHEET)

    Set trans = LLTranslations.Create(shtransLL, shtransDict)
    Set Trads = trans
End Property

Public Property Get geoObject() As ILLGeo
    Dim sh As Worksheet
    Dim geo As ILLGeo
    Set sh = Wkb.Worksheets(GEOSHEET)
    Set geo = LLGeo.Create(sh)
    Set geoObject = geo
End Property

'Test if the translation form exists
Public Property Get TransObject(Optional ByVal trans As TradScopes = TranslationOfMessages) As ITranslation
    Set TransObject = Trads.TransObject(trans)
End Property

Public Property Get FormulaDataObject() As IFormulaData
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(FORMULASHEET)
    Set FormulaDataObject = FormulaData.Create(sh)
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
    Case TemporarySheetAnalysis
        TemporarySheetName = "ana_temp__"
    Case TemporarySheetDebug
        TemporarySheetName = "debug__"
    Case TemporarySheetMetadata
        TemporarySheetName = "metadata__"
    Case TemporarySheetList
        TemporarySheetName = "dropdown_lists__"
    Case TemporarySheetImport
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
    Dim outwkb As Workbook

    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim vars As ILLVariables
    Dim geo As ILLGeo 'Geo Object to add admin names in the preparation steps of the dictionary
    Dim dictTrads As ITranslation
    Dim msgTrads As ITranslation 'Translations of messages object used in preserve sheet Names
    Dim counter As Long

    Dim casewhenVars As BetterArray
    Dim casewhenObject As ICaseWhen
    Dim casewhenVariableName As String
    Dim casewhenFormula As String


    Set dict = Dictionary()
    Set choi = Choices()
    Set dictTrads = TransObject(TranslationOfDictionary)
    Set msgTrads = TransObject()
    Set geo = geoObject()

    Set PreservedSheetNames = New BetterArray
    PreservedSheetNames.LowerBound = 1


    'Preserved names for worksheet
    PreservedSheetNames.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                             PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                             EXPORTSHEET, TemporarySheetName(TemporarySingle), TemporarySheetName(TemporarySheetList), _
                             TemporarySheetName(TemporarySheetDebug), TemporarySheetName(TemporarySheetMetadata), _
                             TemporarySheetName(TemporarySheetAnalysis), TemporarySheetName(TemporarySheetImport), _
                             msgTrads.TranslatedValue("LLSHEET_Admin"), msgTrads.TranslatedValue("LLSHEET_Analysis"), _
                             msgTrads.TranslatedValue("LLSHEET_SpatialAnalysis"), msgTrads.TranslatedValue("LLSHEET_TemporalAnalysis")

    'Translate the dictionary
    dict.Translate dictTrads

    'Prepare the dictionary
    dict.Prepare PreservedSheetNames, geo, tablePrefix:="table"
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
    choi.Translate dictTrads

    'Sort the exports
    choi.Sort
End Sub

'Errors ----------------------------------

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

Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
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

Public Property Get ILinelistSpecs_TransObject(Optional ByVal trans As TradScopes = TranslationOfMessages) As ITranslation
    Set ILinelistSpecs_TransObject = TransObject(trans)
End Property


'Formula data object
Public Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = FormulaDataObject()
End Property

Public Property Get ILinelistSpecs_TemporarysheetName(ByVal scope As TemporarySheetTypes) As String
    ILinelistSpecs_TemporarysheetName = TemporarySheetName(scope)
End Property

Public Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = DesignFormat()
End Property


