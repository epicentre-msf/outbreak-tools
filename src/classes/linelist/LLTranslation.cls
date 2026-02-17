VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist translation tables manager"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Linelist translation tables manager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILLTranslation

'@class LLTranslation
'@description
'Manages translation ListObjects hosted on a single worksheet and provides
'scoped ITranslationObject accessors, import/export of full translation
'tables, dictionary-specific exchange, and workbook-level hidden name
'initialisation for translated configuration values.
'@depends ITranslationObject, TranslationObject, ICustomTable, CustomTable
'@depends IHiddenNames, HiddenNames, BetterArray
'@version 1.0 (2026-02-10)



'@section State
'===============================================================================
'@description Backing UDT and module-level constants.

Private Type TLLTranslation
    translationSheet As Worksheet
End Type

Private this As TLLTranslation

Private Const TRADLLMSG              As String = "T_TradLLMsg"
Private Const TRADLLSHAPES           As String = "T_TradLLShapes"
Private Const TRADLLFORMS            As String = "T_TradLLForms"
Private Const TRADLLRIBBON           As String = "T_TradLLRibbon"
Private Const DICTTRANS              As String = "Tab_Translations"
Private Const CLASS_NAME             As String = "LLTranslation"


'@section Factory
'===============================================================================
'@description Instance creation via VB_PredeclaredId factory pattern.

'@label:create
'@sub-title Create a new LLTranslation from a translation worksheet
'@details
'Validates that the worksheet is not Nothing and that all required
'ListObjects are present (T_TradLLMsg, T_TradLLShapes, T_TradLLForms,
'Tab_Translations). T_TradLLRibbon is optional.
'@param translationSheet Worksheet. The worksheet hosting all translation
'ListObjects.
'@return ILLTranslation. A fully initialised facade.
'@throws ProjectError.ObjectNotInitialized When translationSheet is Nothing.
'@throws ProjectError.ElementNotFound When a required ListObject is missing.
'@export
Public Function Create(ByVal translationSheet As Worksheet) As ILLTranslation
    CheckRequirements translationSheet

    With New LLTranslation
        Set .InternalSheet = translationSheet
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Interface-cast self reference
'@return ILLTranslation. This instance as its interface type.
Public Property Get Self() As ILLTranslation
    Set Self = Me
End Property

'@label:internal-sheet
'@prop-title Friend setter for the host worksheet
'@details
'Used by the factory method during construction.
'@param value Worksheet. The translation worksheet to bind.
Friend Property Set InternalSheet(ByVal value As Worksheet)
    Set this.translationSheet = value
End Property


'@section Properties
'===============================================================================
'@description Accessors for translation objects and the host worksheet.

'@label:wksh
'@prop-title Host translation worksheet
'@details
'Returns the worksheet bound to this instance.
'@return Worksheet. The host translation worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.translationSheet
End Property

'@label:trans-object
'@prop-title Scoped translation object
'@details
'Returns an ITranslationObject configured for the requested translation
'scope. The language code is resolved from the host workbook HiddenNames
'(RNG_LLLanguageCode or RNG_DictionaryLanguage depending on scope).
'@param trans Optional Byte. TranslationScope value selecting the content
'scope. Defaults to TranslationOfMessages.
'@return ITranslationObject. The translation object for the requested scope.
'@throws ProjectError.ElementNotFound When the requested ListObject does
'not exist on the host worksheet.
'@export
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Dim Lo As ListObject
    Dim lang As String
    Dim sh As Worksheet
    Dim wkbNames As IHiddenNames

    Set sh = this.translationSheet
    Set wkbNames = HiddenNames.Create(sh.Parent)
    lang = wkbNames.ValueAsString("RNG_LLLanguageCode")

    Select Case trans
    Case TranslationOfDictionary
        lang = wkbNames.ValueAsString("RNG_DictionaryLanguage")
        Set Lo = ResolveListObject(DICTTRANS)
    Case TranslationOfShapes
        Set Lo = ResolveListObject(TRADLLSHAPES)
    Case TranslationOfForms
        Set Lo = ResolveListObject(TRADLLFORMS)
    Case TranslationOfMessages
        Set Lo = ResolveListObject(TRADLLMSG)
    Case TranslationOfRibbon
        Set Lo = ResolveListObject(TRADLLRIBBON)
    Case Else
        ThrowError ProjectError.InvalidArgument, _
                   "Unknown translation scope: " & CStr(trans)
    End Select

    Set TransObject = TranslationObject.Create(Lo, lang)
End Property


'@section DataExchange
'===============================================================================
'@description Import, export, and dictionary exchange operations.

'@label:export
'@sub-title Export all translation tables to an output workbook
'@details
'Creates or clears a worksheet in the output workbook matching the host
'sheet name, then exports every ListObject from the host sheet using
'CustomTable.Export with stacking. Applies basic formatting and sets
'the sheet visibility to the Hide parameter value.
'@param outputWkb Workbook. The destination workbook.
'@param Hide Optional Long. Sheet visibility after export. Defaults to
'xlSheetVeryHidden.
'@throws ProjectError.ObjectNotInitialized When outputWkb is Nothing.
'@export
Public Sub Export(ByVal outputWkb As Workbook, _
                  Optional ByVal Hide As Long = xlSheetVeryHidden)
    Dim sh As Worksheet
    Dim exportSheet As Worksheet
    Dim Lo As ListObject
    Dim tableAdapter As ICustomTable
    Dim startLine As Long

    If outputWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target workbook is missing; cannot export translations."
    End If

    Set sh = this.translationSheet
    Set exportSheet = FindOrCreateSheet(outputWkb, sh.Name)

    startLine = 1

    For Each Lo In sh.ListObjects
        Set tableAdapter = CustomTable.Create(Lo)
        tableAdapter.Export exportSheet, startLine:=startLine, addListObject:=True, clearSheet:=False
        startLine = startLine + Lo.Range.Rows.Count + 2
    Next

    FormatExportSheet exportSheet, Hide
End Sub

'@label:import
'@sub-title Import translation tables from another workbook
'@details
'Finds a worksheet in the source workbook matching the host sheet name.
'Cleans the dictionary ListObject before import. For each ListObject on
'the host sheet, imports data from the matching ListObject on the source
'sheet. Missing source tables are silently skipped.
'@param fromWkb Workbook. The source workbook containing translation data.
'@export
Public Sub Import(ByVal fromWkb As Workbook)
    Dim sh As Worksheet
    Dim importSheet As Worksheet
    Dim Lo As ListObject
    Dim impLo As ListObject
    Dim tableAdapter As ICustomTable
    Dim impTableAdapter As ICustomTable

    If fromWkb Is Nothing Then Exit Sub

    Set sh = this.translationSheet

    On Error Resume Next
    Set importSheet = fromWkb.Worksheets(sh.Name)
    On Error GoTo 0

    If importSheet Is Nothing Then Exit Sub

    'Clean dictionary table before import
    CleanDictionary

    'Import each matching ListObject
    For Each Lo In sh.ListObjects
        Set impLo = Nothing

        On Error Resume Next
        Set impLo = importSheet.ListObjects(Lo.Name)
        On Error GoTo 0

        If Not impLo Is Nothing Then
            Set tableAdapter = CustomTable.Create(Lo)
            Set impTableAdapter = CustomTable.Create(impLo)
            tableAdapter.Import impTableAdapter
        End If
    Next
End Sub

'@label:import-dictionary
'@sub-title Import dictionary translations from a raw ListObject
'@details
'Overwrites the host dictionary table with data from the supplied
'ListObject, preserving the source column headers via keepSourceHeaders.
'@param dictionaryTable ListObject. The source dictionary ListObject.
'@throws ProjectError.ObjectNotInitialized When dictionaryTable is Nothing.
'@throws ProjectError.ElementNotFound When the host dictionary table is
'not found.
'@export
Public Sub ImportDictionary(ByVal dictionaryTable As ListObject)
    Dim dictLo As ListObject
    Dim tableAdapter As ICustomTable
    Dim impTableAdapter As ICustomTable

    If dictionaryTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary ListObject is required for import."
    End If

    Set dictLo = ResolveListObject(DICTTRANS)
    Set tableAdapter = CustomTable.Create(dictLo)
    Set impTableAdapter = CustomTable.Create(dictionaryTable)
    tableAdapter.Import impTableAdapter, keepSourceHeaders:=True
End Sub

'@label:export-dictionary
'@sub-title Export the dictionary table to a named worksheet
'@details
'Creates or clears a worksheet in the output workbook with the given
'name, then exports the dictionary ListObject using CustomTable.Export.
'Applies formatting and sets the sheet visibility.
'@param outputWkb Workbook. The destination workbook.
'@param sheetName Optional String. Worksheet name for the export.
'Defaults to "Translations".
'@param Hide Optional Long. Sheet visibility after export. Defaults to
'xlSheetVisible.
'@throws ProjectError.ObjectNotInitialized When outputWkb is Nothing.
'@throws ProjectError.ElementNotFound When the host dictionary table is
'not found.
'@export
Public Sub ExportDictionary(ByVal outputWkb As Workbook, _
                            Optional ByVal sheetName As String = "Translations", _
                            Optional ByVal Hide As Long = xlSheetVisible)
    Dim dictLo As ListObject
    Dim tableAdapter As ICustomTable
    Dim exportSheet As Worksheet

    If outputWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target workbook is missing; cannot export dictionary."
    End If

    Set dictLo = ResolveListObject(DICTTRANS)
    Set exportSheet = FindOrCreateSheet(outputWkb, sheetName)
    Set tableAdapter = CustomTable.Create(dictLo)

    tableAdapter.Export exportSheet, startLine:=1, startColumn:=1, addListObject:=True

    FormatExportSheet exportSheet, Hide
End Sub


'@section HiddenNames
'===============================================================================
'@description Workbook-level hidden name initialisation from translations.

'@label:initialise-hidden-names
'@sub-title Create RNG_* hidden names on a target workbook
'@details
'Translates message and sheet-name tags using the TranslationOfMessages
'scope, then writes each translated value as a workbook-level hidden name
'on the target workbook. The RNG_LLLanguage, RNG_LLLanguageCode, and
'RNG_DictionaryLanguage names are NOT created here as they are expected
'to already exist in the designer workbook HiddenNames.
'@param targetWkb Workbook. The workbook receiving the hidden names.
'@throws ProjectError.ObjectNotInitialized When targetWkb is Nothing.
'@export
Public Sub InitialiseHiddenNames(ByVal targetWkb As Workbook)
    Dim wkbNames As IHiddenNames
    Dim msgTrans As ITranslationObject

    If targetWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target workbook is required to initialise hidden names."
    End If

    Set wkbNames = HiddenNames.Create(targetWkb)
    Set msgTrans = TransObject(TranslationOfMessages)

    EnsureAndSet wkbNames, "RNG_GoToSection", msgTrans.TranslatedValue("MSG_GoToSection")
    EnsureAndSet wkbNames, "RNG_AnaPeriod", msgTrans.TranslatedValue("MSG_AnaPeriod")
    EnsureAndSet wkbNames, "RNG_GoToHeader", msgTrans.TranslatedValue("MSG_GoToHead")
    EnsureAndSet wkbNames, "RNG_NoDevide", msgTrans.TranslatedValue("MSG_NoDevide")
    EnsureAndSet wkbNames, "RNG_Devide", msgTrans.TranslatedValue("MSG_Devide")
    EnsureAndSet wkbNames, "RNG_GoToGraph", msgTrans.TranslatedValue("MSG_GoToGraph")
    EnsureAndSet wkbNames, "RNG_OnFiltered", msgTrans.TranslatedValue("MSG_ComputeOnFiltered")
    EnsureAndSet wkbNames, "RNG_CustomDrop", msgTrans.TranslatedValue("LLSHEET_CustomChoice")
    EnsureAndSet wkbNames, "RNG_UASheet", msgTrans.TranslatedValue("LLSHEET_Analysis")
    EnsureAndSet wkbNames, "RNG_TSSheet", msgTrans.TranslatedValue("LLSHEET_TemporalAnalysis")
    EnsureAndSet wkbNames, "RNG_SPSheet", msgTrans.TranslatedValue("LLSHEET_SpatialAnalysis")
    EnsureAndSet wkbNames, "RNG_SPTSheet", msgTrans.TranslatedValue("LLSHEET_SpatioTemporalAnalysis")
    EnsureAndSet wkbNames, "RNG_CustomPivot", msgTrans.TranslatedValue("LLSHEET_CustomPivotTable")
    EnsureAndSet wkbNames, "RNG_Week", msgTrans.TranslatedValue("MSG_W")
    EnsureAndSet wkbNames, "RNG_Quarter", msgTrans.TranslatedValue("MSG_Q")
    EnsureAndSet wkbNames, "RNG_InfoStart", msgTrans.TranslatedValue("MSG_InfoStart")
    EnsureAndSet wkbNames, "RNG_InfoEnd", msgTrans.TranslatedValue("MSG_InfoEnd")
End Sub


'@section SheetNames
'===============================================================================
'@description Worksheet-level HiddenNames for language codes.

'@label:initialise-sheet-names
'@sub-title Create language HiddenNames on the translation worksheet
'@details
'Creates worksheet-level HiddenNames for dictionary and interface language
'codes on the translation worksheet itself. These are later read by
'InitialiseWorkbookNames in Linelist.cls to propagate language settings
'to the output workbook.
'@param dictLang String. The dictionary language code (e.g., "en").
'@param llLang String. The interface language code (e.g., "fr").
'@export
Public Sub InitialiseSheetNames(ByVal dictLang As String, ByVal llLang As String)
    Dim shNames As IHiddenNames

    Set shNames = HiddenNames.Create(this.translationSheet)
    EnsureAndSet shNames, "RNG_DictionaryLanguage", dictLang
    EnsureAndSet shNames, "RNG_LLLanguageCode", llLang
End Sub


'@section ValueAccess
'===============================================================================
'@description Mapped value retrieval from worksheet-level HiddenNames.

'@label:value
'@prop-title Retrieve a named configuration value from the translation worksheet
'@details
'Returns a worksheet-level HiddenName value identified by a logical key.
'Supported keys:
'  "dictlang" -> RNG_DictionaryLanguage
'  "lllang"   -> RNG_LLLanguageCode
'Other keys return vbNullString.
'@param nameId String. The logical key to look up.
'@return String. The stored value, or vbNullString if the key is unknown.
'@export
Public Property Get Value(ByVal nameId As String) As String
    Dim shNames As IHiddenNames
    Dim rngName As String

    Select Case LCase$(Trim$(nameId))
        Case "dictlang"
            rngName = "RNG_DictionaryLanguage"
        Case "lllang"
            rngName = "RNG_LLLanguageCode"
        Case Else
            Exit Property
    End Select

    Set shNames = HiddenNames.Create(this.translationSheet)
    Value = shNames.ValueAsString(rngName)
End Property


'@section Internal Helpers
'===============================================================================
'@description Private utility methods.

'@label:resolve-list-object
'@sub-title Resolve a ListObject by name from the host worksheet
'@details
'Returns the ListObject with the given name from the host translation
'worksheet. Raises an error when the table is not found.
'@param tableName String. The ListObject name to resolve.
'@return ListObject. The resolved table.
'@throws ProjectError.ElementNotFound When the table is not found.
Private Function ResolveListObject(ByVal tableName As String) As ListObject
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = this.translationSheet.ListObjects(tableName)
    On Error GoTo 0

    If Lo Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Table '" & tableName & "' not found on worksheet '" & _
                   this.translationSheet.Name & "'"
    End If

    Set ResolveListObject = Lo
End Function

'@label:find-or-create-sheet
'@sub-title Find or create a worksheet in the target workbook
'@details
'Attempts to find a worksheet with the given name in the target workbook.
'If found, clears its contents entirely. If not found, creates a new
'worksheet with the name appended at the end.
'@param targetWkb Workbook. The target workbook.
'@param targetName String. The worksheet name to find or create.
'@return Worksheet. The resolved or newly created worksheet.
Private Function FindOrCreateSheet(ByVal targetWkb As Workbook, _
                                   ByVal targetName As String) As Worksheet
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = targetWkb.Worksheets(targetName)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = targetWkb.Worksheets.Add( _
            After:=targetWkb.Worksheets(targetWkb.Worksheets.Count))
        sh.Name = targetName
    Else
        sh.Cells.Clear
    End If

    Set FindOrCreateSheet = sh
End Function

'@label:format-export-sheet
'@sub-title Apply standard formatting to an exported sheet
'@details
'Sets font, alignment, autofit, and visibility on the exported sheet.
'@param sh Worksheet. The sheet to format.
'@param Hide Long. Sheet visibility constant.
Private Sub FormatExportSheet(ByVal sh As Worksheet, ByVal Hide As Long)
    On Error Resume Next
    sh.Rows(1).Font.Bold = True
    sh.Cells.Font.Name = "Consolas"
    sh.Cells.HorizontalAlignment = xlHAlignLeft
    sh.Cells.VerticalAlignment = xlVAlignCenter
    sh.Cells.EntireRow.AutoFit
    sh.Cells.EntireColumn.AutoFit
    sh.Visible = Hide
    On Error GoTo 0
End Sub

'@label:clean-dictionary
'@sub-title Clear the dictionary ListObject data
'@details
'Wraps the host dictionary table in a CustomTable and calls Clean to
'clear all non-formula, non-key columns. Silently exits when the
'dictionary table does not exist.
Private Sub CleanDictionary()
    Dim dictLo As ListObject
    Dim tableAdapter As ICustomTable

    On Error Resume Next
    Set dictLo = this.translationSheet.ListObjects(DICTTRANS)
    On Error GoTo 0

    If Not dictLo Is Nothing Then
        Set tableAdapter = CustomTable.Create(dictLo)
        tableAdapter.Clean
    End If
End Sub

'@label:check-requirements
'@sub-title Validate that the translation sheet has all required tables
'@details
'Ensures the worksheet is not Nothing and that the required ListObjects
'(T_TradLLMsg, T_TradLLShapes, T_TradLLForms, Tab_Translations) are
'present. T_TradLLRibbon is optional and not validated.
'@param translationSheet Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When the sheet is Nothing.
'@throws ProjectError.ElementNotFound When a required table is missing.
Private Sub CheckRequirements(ByVal translationSheet As Worksheet)
    Dim requiredTables As BetterArray
    Dim counter As Long
    Dim tableName As String
    Dim Lo As ListObject

    If translationSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Translation sheet is required."
    End If

    Set requiredTables = New BetterArray
    requiredTables.Push TRADLLMSG, TRADLLSHAPES, TRADLLFORMS, DICTTRANS

    For counter = requiredTables.LowerBound To requiredTables.UpperBound
        tableName = requiredTables.Item(counter)
        Set Lo = Nothing

        On Error Resume Next
        Set Lo = translationSheet.ListObjects(tableName)
        On Error GoTo 0

        If Lo Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                       "Required table '" & tableName & "' not found on worksheet '" & _
                       translationSheet.Name & "'"
        End If
    Next
End Sub


'@section Error Handling
'===============================================================================
'@description Centralised error raising.

'@label:ensure-and-set
'@sub-title Create a HiddenName if missing, then always update its value
'@param store IHiddenNames. The hidden names manager.
'@param nameId String. The hidden name identifier.
'@param value String. The value to set.
Private Sub EnsureAndSet(ByVal store As IHiddenNames, ByVal nameId As String, _
                         ByVal value As String)
    store.EnsureName nameId, value, HiddenNameTypeString
    store.SetValue nameId, value
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber ProjectError. The error constant.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ILLTranslation contract.

Private Property Get ILLTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set ILLTranslation_TransObject = TransObject(trans)
End Property

Private Property Get ILLTranslation_Wksh() As Worksheet
    Set ILLTranslation_Wksh = Wksh
End Property

Private Sub ILLTranslation_Export(ByVal outputWkb As Workbook, _
                                  Optional ByVal Hide As Long = xlSheetVeryHidden)
    Export outputWkb, Hide
End Sub

Private Sub ILLTranslation_Import(ByVal fromWkb As Workbook)
    Import fromWkb
End Sub

Private Sub ILLTranslation_ImportDictionary(ByVal dictionaryTable As ListObject)
    ImportDictionary dictionaryTable
End Sub

Private Sub ILLTranslation_ExportDictionary(ByVal outputWkb As Workbook, _
                                            Optional ByVal sheetName As String = "Translations", _
                                            Optional ByVal Hide As Long = xlSheetVisible)
    ExportDictionary outputWkb, sheetName, Hide
End Sub

Private Sub ILLTranslation_InitialiseHiddenNames(ByVal targetWkb As Workbook)
    InitialiseHiddenNames targetWkb
End Sub

Private Sub ILLTranslation_InitialiseSheetNames(ByVal dictLang As String, ByVal llLang As String)
    InitialiseSheetNames dictLang, llLang
End Sub

Private Property Get ILLTranslation_Value(ByVal nameId As String) As String
    ILLTranslation_Value = Value(nameId)
End Property
