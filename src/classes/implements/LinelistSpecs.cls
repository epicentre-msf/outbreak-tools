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
'@IgnoreModule UseMeaningfulName, HungarianNotation
Implements ILinelistSpecs

Option Explicit

Public Enum TemporarySheetTypes
    TemporarySheetAnalysis = 1
    TemporarySheetUpdates = 2
    TemporarySheetMetadata = 3
    TemporarySheetList = 4
    TemporarySheetImport = 5
    TemporarySingle = 6
    TemporarySpatial = 7
    TemporaryShowHide = 8
End Enum

Private Type TLinelistSpecs
    specswkb As Workbook
End Type

Private this As TLinelistSpecs
Private Const CLASSNAME                 As String = "LinelistSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const CHOICESSHEET              As String = "Choices"
Private Const GEOSHEET                  As String = "Geo"
Private Const PASSWORDSHEET             As String = "__pass" 'password sheets
Private Const FORMULASHEET              As String = "__formula" 'sheet with excel formula
Private Const DICTTRANSLATIONSHEET      As String = "Translations" 'Dictionary translation sheet
Private Const HLISTTRANSLATIONSHEET     As String = "LinelistTranslation" 'linelist translation (forms, shapes, etc.)
Private Const ANALYSISSPECSSHEET        As String = "Analysis"
Private Const EXPORTSHEET               As String = "Exports"
Private Const FORMATSHEET               As String = "LinelistStyle"
Private Const MAINSHEET                 As String = "Main"
Private Const DESTRANSSHEET             As String = "DesignerTranslation"

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
    Set dict = LLdictionary.Create(sh, 1, 1, NumberOfExports())
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

Public Property Get DesTrads() As IDesTranslation
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(DESTRANSSHEET)
    Set DesTrads = DesTranslation.Create(sh)
End Property

Public Property Get GeoObject() As ILLGeo
    Dim sh As Worksheet
    Dim geo As ILLGeo
    Set sh = Wkb.Worksheets(GEOSHEET)
    Set geo = LLGeo.Create(sh)
    Set GeoObject = geo
End Property

Public Property Get ExportObject() As ILLExport
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(EXPORTSHEET)
    Set ExportObject = LLExport.Create(sh)
End Property

Private Property Get NumberOfExports() As Long
    NumberOfExports = ExportObject.NumberOfExports()
End Property

Public Property Get Analysis() As ILLAnalysis
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(ANALYSISSPECSSHEET)
    Set Analysis = LLAnalysis.Create(sh)
End Property

Public Property Get Password() As ILLPasswords
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(PASSWORDSHEET)
    Set Password = LLPasswords.Create(sh)
End Property

'Test if the translation form exists
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslation
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

Public Property Get MainObject() As IMain
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(MAINSHEET)
    Set MainObject = Main.Create(sh)
End Property

Public Property Get TemporarySheetName(ByVal scope As TemporarySheetTypes) As String

    Select Case scope
    Case TemporarySheetAnalysis
        TemporarySheetName = "ana_tabnames__"
    Case TemporarySheetUpdates
        TemporarySheetName = "updates__" 'record informations on some updates
    Case TemporarySheetMetadata
        TemporarySheetName = "Metadata"
    Case TemporarySheetList
        TemporarySheetName = "dropdown_lists__" 'Add dropdowns on the linelist
    Case TemporarySheetImport
        TemporarySheetName = "import_rep__" 'for import report
    Case TemporarySingle
        TemporarySheetName = "temp__" 'temporary boggy sheet for future computations
    Case TemporaryShowHide
        TemporarySheetName = "show_hide__" 'Manage show/hide forms
    Case TemporarySpatial
        TemporarySheetName = "spatial_tables__" 'add tables for spatial analysis
    End Select
End Property

'Get Categories for one variable
Private Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray

    Dim categoryData As BetterArray
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim categoryName As String
    Dim categoryType As String
    Dim controlDetails As String

    Set dict = Dictionary()
    Set choi = Choices()
    Set categoryData = New BetterArray
    Set vars = LLVariables.Create(dict)

    categoryType = vars.Value(varName:=varName, colName:="control")
    'This is for multiple choices
    On Error Resume Next
    categoryType = Application.WorksheetFunction.Trim(Split(categoryType, "(")(0))
    On Error GoTo 0
    controlDetails = Application.WorksheetFunction.Trim(vars.Value(varName:=varName, colName:="control details"))

    If categoryType = "choice_formula" Then
        categoryName = Replace(Split(controlDetails, ",")(0), "CHOICE_FORMULA(", "")
    ElseIf (categoryType = "choice_manual") Or _
           (categoryType = "choice_multiple") Or _
           (categoryType = "choice_custom") Then

        categoryName = vars.Value(varName:=varName, colName:="control details")
    End If

    'Test the case_when
    Set categoryData = choi.Categories(categoryName, useShortlabels)
    Set Categories = categoryData.Clone()
End Function

'Speed up before a work
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Return previous state
Private Sub ReturnState()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

Private Sub Prepare()

    'Start Row of the dictionary

    Dim PreservedSheetNames As BetterArray

    Dim dict As ILLdictionary
    Dim impDict As ILLdictionary 'Imported dictionary for checkings
    Dim choi As ILLchoice
    Dim geo As ILLGeo                            'Geo Object to add admin names in the preparation steps of the dictionary
    Dim dictTrads As ITranslation
    Dim msgTrads As ITranslation                 'Translations of messages object used in preserve sheet Names
    Dim mainTrads As IDesTranslation             'Translation of the designer: the main sheet
    Dim mainobj As IMain
    Dim trans As ILLTranslations
    Dim counter As Long
    Dim inpWkb As Workbook
    Dim setPath As String
    Dim sh As Worksheet
    Dim expobj As ILLExport
    Dim ana As ILLAnalysis
    Dim fData As IFormulaData

    Set mainobj = MainObject()
    Set mainTrads = desTrads()

    'If the main object is not ready exit the function
    mainobj.CheckReadiness mainTrads
    If Not mainobj.Ready() Then Exit Sub

    Set geo = geoObject()
        'Import dictionary, and geobase, as well as choices
    setPath = mainobj.ValueOf("setuppath")
    BusyApp
    Set inpWkb = Workbooks.Open(setPath)
    BusyApp

    'All the required sheeets should be present
    sheetExists inpWkb, DICTIONARYSHEET
    sheetExists inpWkb, CHOICESSHEET
    sheetExists inpWkb, DICTTRANSLATIONSHEET
    sheetExists inpWkb, EXPORTSHEET
    sheetExists inpWkb, ANALYSISSPECSSHEET

    'Import the Exports specifications from the setup (Number of exports is used in the dictionary)
    Set expobj = ExportObject()
    Set sh = inpWkb.Worksheets(EXPORTSHEET)
    expobj.ImportSpecs sh, 4, 1

    'Import the dictionary data
    Set dict = Dictionary()    
    Set sh = inpWkb.Worksheets(DICTIONARYSHEET)

    'Exit if the imported dictionary is completely empty
    Set impDict = LLdictionary.Create(sh, 5, 1)
    
    If (Application.WorksheetFunction.COUNTA(impDict.DataRange) = 0) Then
        inpWkb.Close savechanges:=False 
        ThrowError "The dictionary is empty"
    End If
    
    'Import the dicitonary in the current designer
    dict.Import sh, fromStartRow:=5, fromStartcol:=1

    'Import the choice data
    Set sh = inpWkb.Worksheets(CHOICESSHEET)
    Set choi = Choices()
    choi.Import sh, fromStartRow:=4, fromStartcol:=1

    'Import the translations
    Set trans = Trads()
    trans.ImportTranslation inpWkb

    'Import the analysis
    Set sh = inpWkb.Worksheets(ANALYSISSPECSSHEET)
    Set ana = Analysis()
    ana.Import sh

    inpWkb.Close savechanges:=False

    'Import the geobase
    setPath = mainobj.ValueOf("geopath")

    'If the geopath is not empty, then import the geobase
    If (setPath <> vbNullString) Then
        BusyApp
        Set inpWkb = Workbooks.Open(setPath)
        BusyApp

        geo.Import inpWkb
        inpWkb.Close savechanges:=False
    End If

    'Define those objects after the imports
    Set dictTrads = TransObject(TranslationOfDictionary)
    Set msgTrads = TransObject()

    'Preserved names for worksheet
    Set PreservedSheetNames = New BetterArray
    PreservedSheetNames.LowerBound = 1
    PreservedSheetNames.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                             PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, _
                             HLISTTRANSLATIONSHEET, _
                             EXPORTSHEET, _
                             TemporarySheetName(TemporarySingle), _
                             TemporarySheetName(TemporarySheetList), _
                             TemporarySheetName(TemporarySheetUpdates), _
                             TemporarySheetName(TemporarySheetMetadata), _
                             TemporarySheetName(TemporarySheetAnalysis), _
                             TemporarySheetName(TemporarySheetImport), _
                             TemporarySheetName(TemporarySpatial), _
                             TemporarySheetName(TemporaryShowHide), _
                             msgTrads.TranslatedValue("LLSHEET_Admin"), _
                             msgTrads.TranslatedValue("LLSHEET_Analysis"), _
                             msgTrads.TranslatedValue("LLSHEET_SpatialAnalysis"), _
                             msgTrads.TranslatedValue("LLSHEET_TemporalAnalysis"), _
                             msgTrads.TranslatedValue("LLSHEET_CustomChoice"), _
                             msgTrads.TranslatedValue("LLSHEET_CustomPivotTable"), _
                             msgTrads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis")

    'Prepare the dictionary
    dict.Prepare PreservedSheetNames, geo, msgTrads, tablePrefix:="table"

    'Add list auto to dictionary (in a recursive way)
    If Not dict.ColumnExists("list auto") Then AddListAuto

    'Translate the dictionary
    dict.Translate dictTrads

    'Translate the choices
    choi.Translate dictTrads

    'Sort the chocies
    choi.Sort

    'Translate the exports
    expobj.Translate dictTrads

    'Translate the analysis
    ana.Translate dictTrads
End Sub

'Add list auto origin informations
Private Sub AddListAuto()
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim counter As Long
    Dim listVars As BetterArray
    Dim varName As String
    Dim varControl As String
    Dim dictLength As Long

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)

    dict.AddColumn ("list auto")
    Set listVars = dict.Data.FilterData("control", "list_auto", "control details")

    For counter = listVars.LowerBound To listVars.UpperBound
        varName = listVars.Item(counter)

        'listVars can contains the same variable twice, be sure we need to update with the onEmpty parameter
        vars.SetValue varName:=varName, colName:="list auto", newValue:="list_auto_origin", onEmpty:=True

        'If one list auto var is a formula, all the variables in that formula are list auto origin
        'List auto origin triggers update of list auto

        varControl = vars.Value(varName:=varName, colName:="control")
        If (varControl = "formula") Or (varControl = "case_when") Or (varControl = "choice_formula") Then RecursiveListAuto varName
    Next

End Sub

Private Sub RecursiveListAuto(ByVal varName As String)

    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim controlValue As String
    Dim controlDetailsValue As String
    Dim fData As FormulaData
    Dim varFormObject As IFormulas
    Dim formList As BetterArray
    Dim counter As Long
    Dim formvarName As String
    Dim formVarControl As String

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)

    controlValue = vars.Value(varName:=varName, colName:="control")

    If (controlValue = "case_when" Or controlValue = "formula" Or controlValue = "choice_formula") Then
        Set fData = FormulaDataObject()
        controlDetailsValue = vars.Value(varName:=varName, colName:="control details")

        Set varFormObject = Formulas.Create(dict, fData, controlDetailsValue)

        If varFormObject.Valid(formulaType:="linelist") Then

            Set formList = varFormObject.varLists()

            counter = formList.LowerBound

            Do While counter <= formList.UpperBound

                formvarName = formList.Items(counter)
                formVarControl = vars.Value(varName:=formvarName, colName:="control")

                vars.SetValue varName:=formvarName, colName:="list auto", newValue:="list_auto_origin", onEmpty:=True

                If (formVarControl = "formula" Or formVarControl = "case_when") Then RecursiveListAuto formvarName

                counter = counter + 1
            Loop

        End If
    End If

End Sub

Private Sub Export(ByVal toWkb As Workbook)
    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim geo As ILLGeo                            'Geo Object to add admin names in the preparation steps of the dictionary
    Dim trans As ILLTranslations
    Dim expobj As ILLExport
    Dim passobj As ILLPasswords
    Dim mainObj As IMain

    Set dict = Dictionary()
    Set choi = Choices()
    Set trans = Trads()
    Set expobj = ExportObject()
    Set geo = GeoObject()
    Set passobj = Password()
    Set mainObj = MainObject()

    dict.Export toWkb, Hide:=xlSheetVeryHidden
    choi.Export toWkb, Hide:=xlSheetVeryHidden
    trans.ExportToWkb toWkb
    expobj.ExportSpecs toWkb, Hide:=xlSheetVeryHidden
    geo.ExportToWkb toWkb, mainObj.ValueOf("llname"), mainObj.ValueOf("setuppath")
    passobj.Export toWkb

End Sub

'Errors --------------------------------------------------------------------

'Initialize the array of sheets
Private Sub CheckSheetsRequirements(ByVal specswkb As Workbook)

    Dim counter As Long
    Dim sheetName As String
    Dim sheetsArray As BetterArray

    Set sheetsArray = New BetterArray

    sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                     PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                     EXPORTSHEET, FORMATSHEET, MAINSHEET, DESTRANSSHEET

    For counter = sheetsArray.LowerBound To sheetsArray.UpperBound
        sheetName = sheetsArray.Item(counter)
        sheetExists specswkb, sheetName
    Next
End Sub

'Check if a sheet Exists
Private Sub sheetExists(ByVal specswkb As Workbook, ByVal sheetName As String)
    Dim sh As Worksheet                          'Just try to set the workbook if it fails it is closed
    Dim checkExists As Boolean

    On Error Resume Next
    Set sh = specswkb.Worksheets(sheetName)
    checkExists = (Not sh Is Nothing)
    On Error GoTo 0

    If Not checkExists Then ThrowError " Worksheet " & sheetName & " Not found!"

End Sub

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
Private Property Get ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = Dictionary()
End Property

Private Property Get ILinelistSpecs_Choices() As ILLchoice
    Set ILinelistSpecs_Choices = Choices()
End Property

'Categories of one variable
Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Set ILinelistSpecs_Categories = Categories(varName, useShortlabels)
End Function

Private Sub ILinelistSpecs_Prepare()
    Prepare
End Sub

Private Property Get ILinelistSpecs_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslation
    Set ILinelistSpecs_TransObject = TransObject(trans)
End Property

'Formula data object
Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = FormulaDataObject()
End Property

Private Property Get ILinelistSpecs_TemporarysheetName(ByVal scope As TemporarySheetTypes) As String
    ILinelistSpecs_TemporarysheetName = TemporarySheetName(scope)
End Property

Private Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = DesignFormat()
End Property

Private Property Get ILinelistSpecs_SpecsWkb() As Workbook
    Set ILinelistSpecs_SpecsWkb = Wkb()
End Property

Private Property Get ILinelistSpecs_MainObject() As IMain
    Set ILinelistSpecs_MainObject = MainObject()
End Property

Private Property Get ILinelistSpecs_Password() As ILLPasswords
    Set ILinelistSpecs_Password = Password()
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = geoObject()
End Property

Private Property Get ILinelistSpecs_Analysis() As ILLAnalysis
    Set ILinelistSpecs_Analysis = Analysis()
End Property

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    'Ask the user if he wants to see (Or not) the uncompleted linelist
    MsgBox "An internal Error occured when dealing with the setup with specified informations: " & _
            Chr(10) &  textMessage & Chr(10) & _
            "Aborting linelist creation process.", _
            vbExclamation + vbOkOnly, "ERROR"
End Sub

Private Property Get ILinelistSpecs_NumberOfExports() As Long
    ILinelistSpecs_NumberOfExports = NumberOfExports()
End Property