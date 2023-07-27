VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Linelist")
'@ModuleDescription("Linelist Translations")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ILLTranslations

Private Type TLLTranslations
    lltranssh As Worksheet
    dicttranssh As Worksheet
End Type

Private this As TLLTranslations

'Tables for traductions
Private Const TRADLLMSG              As String = "T_TradLLMsg"
Private Const TRADLLSHAPES           As String = "T_TradLLShapes"
Private Const TRADLLFORMS            As String = "T_TradLLForms"
Private Const TRADLLRIBBON           As String = "T_TradLLRibbon"
Private Const LLLANG                 As String = "T_LLLang"
Private Const DICTTRANS              As String = "Tab_Translations"
Private Const SELECTEDLANG           As String = "T_SelectedLLLanguages"
Private Const CLASSNAME              As String = "LLTranslations"

Public Function Create(ByVal lltranssh As Worksheet, _ 
                       ByVal dicttranssh As Worksheet) As ILLTranslations

    CheckRequirements lltranssh, dicttranssh

    With New LLTranslations
        Set .LLTransSheet = lltranssh
        Set .DictTransSheet = dicttranssh
        Set Create = .Self()
    End With
End Function

Public Property Get Self() As ILLTranslations
    Set Self = Me
End Property

Public Property Get LLTransSheet() As Worksheet
    Set LLTransSheet = this.lltranssh
End Property

Public Property Get DictTransSheet() As Worksheet
    Set DictTransSheet = this.dicttranssh
End Property

Public Property Set LLTransSheet(ByVal sh As Worksheet)
    Set this.lltranssh = sh
End Property

Public Property Set DictTransSheet(ByVal sh As Worksheet)
    Set this.dicttranssh = sh
End Property

Private Property Get TransObject(Optional ByVal trans As Byte = _
                                 TranslationOfMessages) As ITranslation

    Dim Lo As ListObject
    Dim lang As String
    Dim sh As Worksheet

    Set sh = LLTransSheet()
    lang = sh.Range("RNG_LLLanguageCode").Value

    Select Case trans
    Case TranslationOfDictionary
        lang = sh.Range("RNG_DictionaryLanguage").Value 'language changed here to the language sheet
        Set sh = DictTransSheet()                'Sheet changed here to the translations sheet
        Set Lo = sh.ListObjects(DICTTRANS)
    Case TranslationOfShapes
        Set Lo = sh.ListObjects(TRADLLSHAPES)
    Case TranslationOfForms
        Set Lo = sh.ListObjects(TRADLLFORMS)
    Case TranslationOfMessages
        Set Lo = sh.ListObjects(TRADLLMSG)
    Case TranslationOfRibbon
        Set Lo = sh.ListObjects(TRADLLRIBBON)
    End Select

    Set TransObject = Translation.Create(Lo, lang)
End Property

Private Property Get Value(ByVal tagName As String) As String
    Dim sh As Worksheet
    Dim retVal As String

    Set sh = LLTransSheet()
    
    On Error Resume Next
    
    Select Case tagName
    Case "lllanguage"
        retVal = sh.Range("RNG_LLLanguage").Value
    Case "lllanguagecode"
         retVal = sh.Range("RNG_LLLanguageCode").Value
    Case "gotosection"
         retVal = sh.Range("RNG_GoToSection").Value
    Case "anaperiod"
        retVal = sh.Range("RNG_AnaPeriod").Value
    Case "gotoheader"
         retVal = sh.Range("RNG_GoToHeader").Value
    Case "dictionarylanguage"
         retVal = sh.Range("RNG_DictionaryLanguage").Value
    Case "nodevide"
        retVal = sh.Range("RNG_NoDevide").Value
    Case "devide"
         retVal = sh.Range("RNG_Devide").Value
    Case "gotograph"
         retVal = sh.Range("RNG_GoToGraph").Value
    Case "onfiltered"
         retVal = sh.Range("RNG_OnFiltered").Value
    Case "customdropdown"
        retVal = sh.Range("RNG_CustomDrop").Value
    Case "uasheet"
        retVal = sh.Range("RNG_UASheet").Value
    Case "tssheet"
        retVal = sh.Range("RNG_TSSheet").Value
    Case "spsheet"
        retVal = sh.Range("RNG_SPSheet").Value
    Case "sptsheet"
        retVal = sh.Range("RNG_SPTSheet").Value
    Case "customtable"
        retVal = sh.Range("RNG_CustomPivot").Value
    End Select

    On Error GoTo 0

    Value = retVal
End Property

Private Sub ExportToWkb(ByVal outputWkb As Workbook)

    Dim outshLLTrans As Worksheet
    Dim outshDictTrans As Worksheet
    Dim shLL As Worksheet
    Dim shDict As Worksheet
    Dim Lo As ListObject
    Dim outRng As Range
    Dim outRngAddress As String
    Dim rngTable As BetterArray
    Dim tablRng As BetterArray
    Dim counter As Long
    Dim rngName As String

    Set shLL = LLTransSheet()
    Set shDict = DictTransSheet()
    Set tablRng = New BetterArray

    'Test if there is no sheet name translation or Linelistranslation in the output workbook
    On Error Resume Next
    Set outshLLTrans = outputWkb.Worksheets(shLL.Name)
    On Error GoTo 0
    If Not outshLLTrans Is Nothing Then ThrowError "Unable to export, already a sheet named " & shLL.Name & " in the workbook"

    On Error Resume Next
    Set outshDictTrans = outputWkb.Worksheets(shDict.Name)
    On Error GoTo 0
    If Not outshDictTrans Is Nothing Then ThrowError "Unable to export, already a sheet named " & shDict.Name & " in the workbook"

    'Add LinelistTranslation and Dictionary translation worksheets
    outputWkb.Worksheets.Add.Name = shLL.Name
    Set outshLLTrans = outputWkb.Worksheets(shLL.Name)

    outputWkb.Worksheets.Add.Name = shDict.Name
    Set outshDictTrans = outputWkb.Worksheets(shDict.Name)

    'Dictionary translation
    Set Lo = shDict.ListObjects(DICTTRANS)

    'I will start at column 1, line 1 for translations in the linelist
    With outshDictTrans
        Set outRng = .Range(.Cells(1, 1), .Cells(Lo.Range.Rows.Count, Lo.Range.Columns.Count))
        outRng.Value = Lo.Range.Value
        'Define the listobject
        .ListObjects.Add(xlSrcRange, outRng, , xlYes).Name = Lo.Name
        .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
    End With

    'All the listObjects in the linelistTranslation
    For Each Lo In shLL.ListObjects
        outRngAddress = Lo.Range.Address
        Set outRng = outshLLTrans.Range(outRngAddress)
        outRng.Value = Lo.Range.Value
        'Add the listObject
        With outshLLTrans
            .ListObjects.Add(xlSrcRange, outRng, , xlYes).Name = Lo.Name
            .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
        End With
    Next

    'Ranges (For more ranghes we should use a table of the range names)
    tablRng.Push "RNG_LLLanguage", "RNG_LLLanguageCode", "RNG_GoToSection", _
                 "RNG_AnaPeriod", "RNG_GoToHeader", "RNG_DictionaryLanguage", _
                 "RNG_NoDevide", "RNG_Devide", "RNG_GoToGraph", "RNG_OnFiltered", _ 
                 "RNG_CustomDrop", "RNG_UASheet", "RNG_TSSheet", "RNG_SPSheet", _ 
                 "RNG_SPTSheet", "RNG_CustomPivot"

    For counter = tablRng.LowerBound To tablRng.UpperBound
        rngName = tablRng.Item(counter)
        outshLLTrans.Range(shLL.Range(rngName).Address).Name = rngName
        outshLLTrans.Range(rngName).Value = shLL.Range(rngName).Value

        'Add Formula
        On Error Resume Next
        outshLLTrans.Range(rngName).formula = shLL.Range(rngName).formula
        On Error GoTo 0
    Next

    outshDictTrans.Visible = xlSheetVeryHidden
    outshLLTrans.Visible = xlSheetVeryHidden
    outshLLTrans.calculate
End Sub

Private Sub ImportTranslation(ByVal fromWkb As Workbook)

    Dim sh As Worksheet
    Dim impsh As Worksheet
    Dim Lo As ListObject
    Dim rng As Range

    Set sh = DictTransSheet()

    'Test for the worksheet existence
    On Error Resume Next
    Set impsh = fromWkb.Worksheets(sh.Name)
    On Error GoTo 0
    If impsh Is Nothing Then ThrowError "I can't find the sheet " & sh.Name & " for import"

    'Test for the listObject existence in the worksheet
    On Error Resume Next
    Set Lo = impsh.ListObjects(DICTTRANS)
    On Error GoTo 0
    If Lo Is Nothing Then ThrowError " No table " & DICTTRANS & "found in worksheet" & impsh.Name

    With sh
        'Unlist the listobject in the translation worksheet
        On Error Resume Next
        .ListObjects(DICTTRANS).Unlist
        On Error GoTo 0
        .Cells.Clear
        Set rng = .Range(Lo.Range.Address)
        rng.Value = Lo.Range.Value
        'Redefine the list Object
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = DICTTRANS
        .ListObjects(DICTTRANS).TableStyle = Lo.TableStyle
    End With
End Sub


'Requirements for instanciation
Private Sub CheckRequirements(ByVal lltradsh As Worksheet, dtradsh As Worksheet)
    Dim tabl As BetterArray
    Dim tablRng As BetterArray 'List of ranges that should exists in the lltradsh
    Dim counter As Long

    If lltradsh Is Nothing Then ThrowError "Unable to instanciate, no linelist translation sheet"
    If dtradsh Is Nothing Then ThrowError "Unable to instanciate, no dictionary translation sheet"

    'Required listObjects fo translations
    Set tabl = New BetterArray

    tabl.Push TRADLLMSG, TRADLLSHAPES, TRADLLFORMS, LLLANG, SELECTEDLANG
    LoExists tabl, lltradsh

    Set tablRng = New BetterArray

    tablRng.Push "RNG_LLLanguage", "RNG_LLLanguageCode", "RNG_GoToSection", _
                 "RNG_AnaPeriod", "RNG_GoToHeader", "RNG_DictionaryLanguage", _
                 "RNG_NoDevide", "RNG_Devide", "RNG_GoToGraph", "RNG_OnFiltered", _ 
                 "RNG_CustomDrop", "RNG_UASheet", "RNG_TSSheet", "RNG_SPSheet", _ 
                 "RNG_SPTSheet", "RNG_CustomPivot"
    For counter = tablRng.LowerBound to tablRng.UpperBound
        RangeExists rngName:=tablRng.Item(counter), sh:=lltradsh
    Next
End Sub

'Check if listobject exists in a worksheet
Private Sub RangeExists(ByVal rngName As String, ByVal sh As Worksheet)

    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name
End Sub

Private Sub LoExists(ByVal loNames As BetterArray, sh As Worksheet)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String
    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError "Table " & lName & " does not exists in Worksheet " & sh.Name
    Next
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Interfaces
Public Property Get ILLTranslations_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslation
    Set ILLTranslations_TransObject = TransObject(trans)
End Property

Public Sub ILLTranslations_ExportToWkb(ByVal outputWkb As Workbook)
    ExportToWkb outputWkb
End Sub

Public Sub ILLTranslations_ImportTranslation(ByVal fromWkb As Workbook)
    ImportTranslation fromWkb
End Sub

Private Property Get ILLTranslations_Value(ByVal tagName As String) As String
    ILLTranslations_Value = Value(tagName)
End Property
