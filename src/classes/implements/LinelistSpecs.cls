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

Public Property Get trads() As ILLTranslations
    Dim trans As ILLTranslations
    Dim shtransLL As Worksheet
    Dim shtransDict As Worksheet

    Set shtransLL = Wkb.Worksheets(HLISTTRANSLATIONSHEET)
    Set shtransDict = Wkb.Worksheets(DICTTRANSLATIONSHEET)

    Set trans = LLTranslations.Create(shtransLL, shtransDict)
    Set trads = trans
End Property

Public Property Get desTrads() As IDesTranslation
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(DESTRANSSHEET)
    Set desTrads = DesTranslation.Create(sh)
End Property

Public Property Get geoObject() As ILLGeo
    Dim sh As Worksheet
    Dim geo As ILLGeo
    Set sh = Wkb.Worksheets(GEOSHEET)
    Set geo = LLGeo.Create(sh)
    Set geoObject = geo
End Property

Public Property Get Exports() As ILLExport
    Dim sh As Worksheet
    Set sh = Wkb.Worksheets(EXPORTSHEET)
    Set Exports = LLExport.Create(sh)
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
    Set TransObject = trads.TransObject(trans)
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
Private Function Categories(ByVal varName As String) As BetterArray

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
    controlDetails = Application.WorksheetFunction.Trim(vars.Value(varName:=varName, colName:="control details"))

    If categoryType = "choice_formula" Then
        categoryName = Replace(Split(controlDetails, ",")(0), "CHOICE_FORMULA(", "")
    ElseIf categoryType = "choice_manual" Then

        categoryName = vars.Value(varName:=varName, colName:="control details")
        'Normally there is no more case_when category, but I keep it
    ElseIf categoryType = "case_when" Then
        categoryName = "__case_when_" & varName
    End If

    'Test the case_when
    Set categoryData = choi.Categories(categoryName)
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

    Set dict = Dictionary()
    Set choi = Choices()
    Set trans = trads()
    Set expobj = Exports()
    Set ana = Analysis()
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

    'Import the dictionary data
    Set sh = inpWkb.Worksheets(DICTIONARYSHEET)
    dict.Import sh, fromStartRow:=5, fromStartcol:=1

    'Import the choice data
    Set sh = inpWkb.Worksheets(CHOICESSHEET)
    choi.Import sh, fromStartRow:=4, fromStartcol:=1

    'Import the translations
    trans.ImportTranslation inpWkb

    'Import the Exports specifications from the setup
    Set sh = inpWkb.Worksheets(EXPORTSHEET)
    expobj.ImportSpecs sh, 4, 1

    'Import the analysis
    Set sh = inpWkb.Worksheets(ANALYSISSPECSSHEET)
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
                             PASSWORDSHEET, FORMULASHEET, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                             EXPORTSHEET, TemporarySheetName(TemporarySingle), TemporarySheetName(TemporarySheetList), _
                             TemporarySheetName(TemporarySheetUpdates), TemporarySheetName(TemporarySheetMetadata), _
                             TemporarySheetName(TemporarySheetAnalysis), TemporarySheetName(TemporarySheetImport), _
                             TemporarySheetName(TemporarySpatial), TemporarySheetName(TemporaryShowHide), _
                             msgTrads.TranslatedValue("LLSHEET_Admin"), msgTrads.TranslatedValue("LLSHEET_Analysis"), _
                             msgTrads.TranslatedValue("LLSHEET_SpatialAnalysis"), msgTrads.TranslatedValue("LLSHEET_TemporalAnalysis"), _
                             msgTrads.TranslatedValue("LLSHEET_CustomChoice")

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

    mainobj.UpdateStatus 10

    'Return the state of the application
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

            Set formList = varFormObject.Varlists()

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

    Set dict = Dictionary()
    Set choi = Choices()
    Set trans = trads()
    Set expobj = Exports()
    Set geo = geoObject()
    Set passobj = Password()

    dict.export toWkb
    choi.export toWkb, Hide:=xlSheetVeryHidden
    trans.ExportToWkb toWkb
    expobj.ExportSpecs toWkb, Hide:=xlSheetVeryHidden
    geo.ExportToWkb toWkb
    passobj.export toWkb

End Sub

'Errors --------------------------------------------------------------------

'Initialize the array of sheets
Private Sub CheckSheetsRequirements(ByVal specswkb As Workbook)

    Dim counter As Long
    Dim sheetName As String
    Dim sheetsArray As BetterArray

    Set sheetsArray = New BetterArray

    sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, GEOSHEET, _
                     PASSWORDSHEET, FormulaSheet, DICTTRANSLATIONSHEET, HLISTTRANSLATIONSHEET, _
                     EXPORTSHEET, FORMATSHEET, MAINSHEET, DESTRANSSHEET

    For counter = sheetsArray.LowerBound To sheetsArray.UpperBound
        sheetName = sheetsArray.Item(counter)
        sheetExists specswkb, sheetName
    Next
End Sub

'Chek if a sheet really exists

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

Public Property Get ILinelistSpecs_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslation
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

Public Property Get ILinelistSpecs_SpecsWkb() As Workbook
    Set ILinelistSpecs_SpecsWkb = Wkb()
End Property

Public Property Get ILinelistSpecs_MainObject() As IMain
    Set ILinelistSpecs_MainObject = MainObject()
End Property

Public Property Get ILinelistSpecs_Password() As ILLPasswords
    Set ILinelistSpecs_Password = Password()
End Property

Public Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = geoObject()
End Property

Public Property Get ILinelistSpecs_Analysis() As ILLAnalysis
    Set ILinelistSpecs_Analysis = Analysis()
End Property

Public Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    export toWkb
End Sub
