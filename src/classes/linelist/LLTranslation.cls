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


'@section Enumerations
'===============================================================================
'@description Supported translation content scopes.

'@sub-title Available translation content scopes
'@details
'Each value maps to a specific ListObject on the translation worksheet.
'Used as a Byte parameter to select which translation table backs the
'returned TranslationObject.
Public Enum TranslationScope
    TranslationOfMessages = 0
    TranslationOfDictionary = 1
    TranslationOfShapes = 2
    TranslationOfForms = 3
    TranslationOfRibbon = 4
End Enum


'@class LLTranslation
'@description
'Manages the translation ListObjects that sit on one worksheet and hands out a
'TranslationObject per scope. TransObject answers the object of one scope,
'Refresh drops what is held, Import and Export move whole translation tables
'between workbooks, ImportDictionary and ExportDictionary move the setup
'translation table on its own, and InitialiseHiddenNames writes the translated
'configuration values into a workbook as hidden names.
'
'WHAT THE FIVE TABLES HOLD
'-------------------------------------------------------------------------------
'Four of them translate the linelist itself: T_TradLLMsg the messages,
'T_TradLLShapes the button captions, T_TradLLForms the form labels and
'T_TradLLRibbon the ribbon labels.
'
'Tab_Translations is the SETUP's translation table, which is a different thing
'from the linelist dictionary. The setup file carries it, InitTransfer copies
'it onto the linelist's copy of this worksheet while the linelist is being
'built, and it is what translates the dictionary, the choices and the analyses.
'ImportDictionary is the copy in and ExportDictionary the copy out: a data
'export from a linelist carries this table on its own.
'
'THE FIVE TABLES ARE ALL REQUIRED
'-------------------------------------------------------------------------------
'Create asks for all five. The designer template ships every one of them, so a
'sheet missing one is a broken sheet and the failure is raised at creation.
'
'WHERE THE LANGUAGES LIVE
'-------------------------------------------------------------------------------
'The two language codes live in workbook-level hidden names,
'RNG_LLLanguageCode and RNG_DictionaryLanguage, on the workbook that holds the
'translation worksheet. TransObject reads them there and raises when the code
'it needs is empty.
'
'WHAT IS HELD IN MEMORY
'-------------------------------------------------------------------------------
'The hidden name store and the five scoped TranslationObjects are built once
'per instance. Building the store walks every Name of the workbook, and each
'TranslationObject holds its own copy of its table, so a caller that keeps this
'instance keeps both. Refresh drops all of them. The two import routines drop
'the scoped objects themselves, and a caller that writes into a translation
'table by any other route calls Refresh.
'@depends TranslationObject, CustomTable, HiddenNames, Checking
'@version 1.2 (2026-07-30)



'@section State
'===============================================================================
'@description Backing UDT and module-level constants.

Private Type TLLTranslation
    translationSheet As Worksheet
    'Building a HiddenNames store walks every Name of its container, and the
    'store was rebuilt on every read.
    wkbNames As HiddenNames
    'One field per scope. TranslationObject holds its lookup per instance, so a
    'caller that gets a fresh object on every call gets none of that copy.
    transMessages As TranslationObject
    transDictionary As TranslationObject
    transShapes As TranslationObject
    transForms As TranslationObject
    transRibbon As TranslationObject
    sealed As Boolean
End Type

Private this As TLLTranslation

Private Const TRADLLMSG              As String = "T_TradLLMsg"
Private Const TRADLLSHAPES           As String = "T_TradLLShapes"
Private Const TRADLLFORMS            As String = "T_TradLLForms"
Private Const TRADLLRIBBON           As String = "T_TradLLRibbon"
Private Const DICTTRANS              As String = "Tab_Translations"
Private Const LANG_INTERFACE         As String = "RNG_LLLanguageCode"
Private Const LANG_DICTIONARY        As String = "RNG_DictionaryLanguage"
Private Const CLASS_NAME             As String = "LLTranslation"


'@section Factory
'===============================================================================
'@description Instance creation via VB_PredeclaredId factory pattern.

'@label:create
'@sub-title Create a new LLTranslation from a translation worksheet
'@details
'Validates that the worksheet is not Nothing and that all five translation
'ListObjects are present (T_TradLLMsg, T_TradLLShapes, T_TradLLForms,
'T_TradLLRibbon, Tab_Translations).
'@param translationSheet Worksheet. The worksheet hosting all translation
'ListObjects.
'@return LLTranslation. A facade over the five tables.
'@throws ProjectError.ObjectNotInitialized When translationSheet is Nothing.
'@throws ProjectError.ElementNotFound When a required ListObject is missing.
'@export
Public Function Create(ByVal translationSheet As Worksheet) As LLTranslation
    CheckRequirements translationSheet

    Dim newTranslation As LLTranslation
    Set newTranslation = New LLTranslation
    Set newTranslation.InternalSheet = translationSheet
    newTranslation.Seal
    Set Create = newTranslation
End Function

'@label:seal
'@sub-title Seal the instance against further creation-only writes
Public Sub Seal()
    this.sealed = True
End Sub

'@label:internal-sheet
'@prop-title Friend setter for the host worksheet
'@details
'Used by the factory method during construction. Writing to it after Create
'raises, so the cached hidden name store and the cached translation objects
'can never point at a worksheet the instance has given up.
'@param value Worksheet. The translation worksheet to bind.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set InternalSheet(ByVal value As Worksheet)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, _
                   "InternalSheet is set at creation only"
    Set this.translationSheet = value
End Property

'@label:workbook-names
'@sub-title The hidden name store of the workbook holding the translations
'@details
'Built once per instance. Building one walks every Name of the workbook.
'@return HiddenNames. The workbook-level store.
Private Function WorkbookNames() As HiddenNames
    If this.wkbNames Is Nothing Then
        Set this.wkbNames = HiddenNames.Create(this.translationSheet.Parent)
    End If
    Set WorkbookNames = this.wkbNames
End Function


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
'Returns a TranslationObject configured for the requested translation scope.
'The language code is resolved from the hidden names of the workbook holding
'the translation worksheet: RNG_DictionaryLanguage for the dictionary scope
'and RNG_LLLanguageCode for the other four.
'
'Each scope is built once and kept. Refresh drops the five.
'@param trans Optional Byte. TranslationScope value selecting the content
'scope. Defaults to TranslationOfMessages.
'@return TranslationObject. The translation object for the requested scope.
'@throws ProjectError.InvalidArgument When the scope is outside the enum.
'@throws ProjectError.ElementNotFound When the requested ListObject is absent
'from the host worksheet, or when the language code the scope needs is empty.
'@export
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As TranslationObject
    Select Case trans
    Case TranslationOfMessages
        If this.transMessages Is Nothing Then _
            Set this.transMessages = BuildScope(TRADLLMSG, LANG_INTERFACE)
        Set TransObject = this.transMessages

    Case TranslationOfDictionary
        If this.transDictionary Is Nothing Then _
            Set this.transDictionary = BuildScope(DICTTRANS, LANG_DICTIONARY)
        Set TransObject = this.transDictionary

    Case TranslationOfShapes
        If this.transShapes Is Nothing Then _
            Set this.transShapes = BuildScope(TRADLLSHAPES, LANG_INTERFACE)
        Set TransObject = this.transShapes

    Case TranslationOfForms
        If this.transForms Is Nothing Then _
            Set this.transForms = BuildScope(TRADLLFORMS, LANG_INTERFACE)
        Set TransObject = this.transForms

    Case TranslationOfRibbon
        If this.transRibbon Is Nothing Then _
            Set this.transRibbon = BuildScope(TRADLLRIBBON, LANG_INTERFACE)
        Set TransObject = this.transRibbon

    Case Else
        ThrowError ProjectError.InvalidArgument, _
                   "Unknown translation scope: " & CStr(trans)
    End Select
End Property

'@label:refresh
'@sub-title Drop everything this instance holds in memory
'@details
'Clears the hidden name store and the five scoped translation objects, so the
'next read builds them again. Call it after a name is deleted and added back
'on the workbook, and after writing into a translation table by any route
'other than this class.
'@export
Public Sub Refresh()
    Set this.wkbNames = Nothing
    DropScopeCache
End Sub


'@section DataExchange
'===============================================================================
'@description Import, export, and dictionary exchange operations.

'@label:export
'@sub-title Export all translation tables to an output workbook
'@details
'Creates or empties a worksheet in the output workbook matching the host
'sheet name, then exports every ListObject from the host sheet using
'CustomTable.Export with stacking. Formats the used range and sets the sheet
'visibility to the Hide parameter value.
'
'Tab_Translations is written LAST, under the other four. InitTransfer
'overwrites it with the setup's own table as soon as the sheet has arrived, and
'CustomTable.Import writes the incoming rows over the cells below the table
'rather than inserting them. Anywhere but the bottom of the stack, a setup
'table with more rows than the designer carries would land on the table below.
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
    Dim tableAdapter As CustomTable
    Dim startLine As Long

    If outputWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target workbook is missing; cannot export translations."
    End If

    Set sh = this.translationSheet
    Set exportSheet = FindOrCreateSheet(outputWkb, sh.Name)

    startLine = 1

    For Each Lo In sh.ListObjects
        If Lo.Name <> DICTTRANS Then
            Set tableAdapter = CustomTable.Create(Lo)
            tableAdapter.Export exportSheet, startLine:=startLine, addListObject:=True, clearSheet:=False
            startLine = startLine + Lo.Range.Rows.Count + 2
        End If
    Next

    'The setup table closes the stack, so the rewrite it takes on the target
    'sheet has empty rows under it to grow into.
    Set Lo = ResolveListObject(DICTTRANS)
    Set tableAdapter = CustomTable.Create(Lo)
    tableAdapter.Export exportSheet, startLine:=startLine, addListObject:=True, clearSheet:=False

    FormatExportSheet exportSheet, Hide
End Sub

'@label:import
'@sub-title Import translation tables from another workbook
'@details
'Finds the worksheet in the source workbook carrying the host sheet name, and
'for each ListObject on the host sheet imports the data of the matching
'ListObject on that source sheet. A table the source does not carry is left
'standing.
'
'Tab_Translations used to be emptied here before the loop had discovered
'whether the source carried a replacement, and it was emptied through a
'CustomTable with no key column, which takes the tag column with the rest. A
'source without Tab_Translations left the setup translations blank with
'nothing to write them back.
'@param fromWkb Workbook. The source workbook containing translation data.
'@throws ProjectError.ObjectNotInitialized When fromWkb is Nothing.
'@throws ProjectError.ElementNotFound When the source workbook carries no
'worksheet of the host sheet name.
'@export
Public Sub Import(ByVal fromWkb As Workbook)
    Dim sh As Worksheet
    Dim importSheet As Worksheet
    Dim Lo As ListObject
    Dim impLo As ListObject
    Dim tableAdapter As CustomTable
    Dim impTableAdapter As CustomTable

    If fromWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Source workbook is missing; cannot import translations."
    End If

    Set sh = this.translationSheet

    On Error Resume Next
    Set importSheet = fromWkb.Worksheets(sh.Name)
    On Error GoTo 0

    If importSheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Worksheet '" & sh.Name & "' not found in workbook '" & _
                   fromWkb.Name & "'; cannot import translations."
    End If

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

    DropScopeCache
End Sub

'@label:import-dictionary
'@sub-title Take the setup translation table onto the host worksheet
'@details
'Overwrites Tab_Translations with the supplied ListObject, which is the
'translation table of the setup file. The host table takes the source column
'headers with keepSourceHeaders, so it comes out carrying the language columns
'the setup carries. InitTransfer calls this on the LINELIST's translation
'sheet while the linelist is being built, and what lands here is what
'translates the dictionary, the choices and the analyses. The designer's own
'translation sheet is never written this way: its four linelist tables are
'maintained by hand, through the designer's import button.
'@param dictionaryTable ListObject. The setup translation ListObject.
'@throws ProjectError.ObjectNotInitialized When dictionaryTable is Nothing.
'@throws ProjectError.ElementNotFound When the host dictionary table is
'not found.
'@export
Public Sub ImportDictionary(ByVal dictionaryTable As ListObject)
    Dim dictLo As ListObject
    Dim tableAdapter As CustomTable
    Dim impTableAdapter As CustomTable

    If dictionaryTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary ListObject is required for import."
    End If

    Set dictLo = ResolveListObject(DICTTRANS)
    Set tableAdapter = CustomTable.Create(dictLo)
    Set impTableAdapter = CustomTable.Create(dictionaryTable)
    tableAdapter.Import impTableAdapter, keepSourceHeaders:=True

    DropScopeCache
End Sub

'@label:export-dictionary
'@sub-title Write the setup translation table onto a worksheet of an output workbook
'@details
'Creates or empties a worksheet in the output workbook with the given
'name, then exports Tab_Translations using CustomTable.Export. Formats the
'used range and sets the sheet visibility.
'
'LLExporter calls this when a user exports data out of a linelist. Of the five
'tables, this is the one that travels: it is what lets the exported dictionary
'and choices still be read in the language the user picked.
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
    Dim tableAdapter As CustomTable
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
'on the target workbook. The three language names are left alone: they are
'written by InitTransfer and by Linelist before this runs.
'
'When the target workbook is the one holding the translation worksheet, the
'store this instance already carries is used, which saves a second walk over
'the same Names collection.
'@param targetWkb Workbook. The workbook receiving the hidden names.
'@throws ProjectError.ObjectNotInitialized When targetWkb is Nothing.
'@export
Public Sub InitialiseHiddenNames(ByVal targetWkb As Workbook)
    Dim targetNames As HiddenNames
    Dim msgTrans As TranslationObject

    If targetWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target workbook is required to initialise hidden names."
    End If

    If targetWkb Is this.translationSheet.Parent Then
        Set targetNames = WorkbookNames()
    Else
        Set targetNames = HiddenNames.Create(targetWkb)
    End If

    Set msgTrans = TransObject(TranslationOfMessages)

    'MSG_GoToSec, the code the dropdowns themselves are built from. AnalysisOutput
    'and LLDataEntry both write their go-to captions from MSG_GoToSec, and the
    'analysis change handler strips this stored caption off the picked entry
    'before it looks the section up. Reading a different code here left the
    'handler stripping a caption no cell carries: a lookup that misses answers
    'with the tag it was handed, so the search ran on "MSG_GoToSection: Demography"
    'and found nothing, on every analysis sheet.
    EnsureAndSet targetNames, "RNG_GoToSection", msgTrans.TranslatedValue("MSG_GoToSec")
    EnsureAndSet targetNames, "RNG_AnaPeriod", msgTrans.TranslatedValue("MSG_AnaPeriod")
    EnsureAndSet targetNames, "RNG_GoToHeader", msgTrans.TranslatedValue("MSG_GoToHead")
    EnsureAndSet targetNames, "RNG_NoDevide", msgTrans.TranslatedValue("MSG_NoDevide")
    EnsureAndSet targetNames, "RNG_Devide", msgTrans.TranslatedValue("MSG_Devide")
    EnsureAndSet targetNames, "RNG_GoToGraph", msgTrans.TranslatedValue("MSG_GoToGraph")
    EnsureAndSet targetNames, "RNG_OnFiltered", msgTrans.TranslatedValue("MSG_ComputeOnFiltered")
    EnsureAndSet targetNames, "RNG_CustomDrop", msgTrans.TranslatedValue("LLSHEET_CustomChoice")
    EnsureAndSet targetNames, "RNG_UASheet", msgTrans.TranslatedValue("LLSHEET_Analysis")
    EnsureAndSet targetNames, "RNG_TSSheet", msgTrans.TranslatedValue("LLSHEET_TemporalAnalysis")
    EnsureAndSet targetNames, "RNG_SPSheet", msgTrans.TranslatedValue("LLSHEET_SpatialAnalysis")
    EnsureAndSet targetNames, "RNG_SPTSheet", msgTrans.TranslatedValue("LLSHEET_SpatioTemporalAnalysis")
    EnsureAndSet targetNames, "RNG_CustomPivot", msgTrans.TranslatedValue("LLSHEET_CustomPivotTable")
    EnsureAndSet targetNames, "RNG_Week", msgTrans.TranslatedValue("MSG_W")
    EnsureAndSet targetNames, "RNG_Quarter", msgTrans.TranslatedValue("MSG_Q")
    EnsureAndSet targetNames, "RNG_InfoStart", msgTrans.TranslatedValue("MSG_InfoStart")
    EnsureAndSet targetNames, "RNG_InfoEnd", msgTrans.TranslatedValue("MSG_InfoEnd")
End Sub


'@section Internal Helpers
'===============================================================================
'@description Private utility methods.

'@label:build-scope
'@sub-title Build the translation object of one scope
'@details
'Resolves the table and the language code the scope needs, and raises when
'the code is empty. An empty code gives a TranslationObject whose language
'matches no column header, and every tag then translates to itself with no
'error raised anywhere below it.
'@param tableName String. The ListObject backing the scope.
'@param langNameId String. The hidden name carrying the language code.
'@return TranslationObject. The translation object for that table.
'@throws ProjectError.ElementNotFound When the table is absent or the
'language code is empty.
Private Function BuildScope(ByVal tableName As String, _
                            ByVal langNameId As String) As TranslationObject
    Dim Lo As ListObject
    Dim lang As String

    Set Lo = ResolveListObject(tableName)
    lang = WorkbookNames().ValueAsString(langNameId)

    If LenB(lang) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The hidden name '" & langNameId & "' holds no language code " & _
                   "in workbook '" & this.translationSheet.Parent.Name & "', so " & _
                   "table '" & tableName & "' cannot be read in any language."
    End If

    Set BuildScope = TranslationObject.Create(Lo, lang)
End Function

'@label:drop-scope-cache
'@sub-title Forget the five scoped translation objects
'@details
'Each one holds a copy of its table, so anything that rewrites a table has to
'drop them.
Private Sub DropScopeCache()
    Set this.transMessages = Nothing
    Set this.transDictionary = Nothing
    Set this.transShapes = Nothing
    Set this.transForms = Nothing
    Set this.transRibbon = Nothing
End Sub

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
'Finds the worksheet of the given name in the target workbook. When it is
'there, its tables are deleted and its cells cleared; when it is absent, a
'worksheet of that name is added at the end.
'
'Clearing cells leaves the ListObject objects standing, and each one keeps
'its name. The export that follows adds tables under those same names, so a
'second export into one workbook used to meet a name Excel had already given
'out and raise 1004.
'@param targetWkb Workbook. The target workbook.
'@param targetName String. The worksheet name to find or create.
'@return Worksheet. The resolved or newly created worksheet.
Private Function FindOrCreateSheet(ByVal targetWkb As Workbook, _
                                   ByVal targetName As String) As Worksheet
    Dim sh As Worksheet
    Dim counter As Long

    On Error Resume Next
    Set sh = targetWkb.Worksheets(targetName)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = targetWkb.Worksheets.Add( _
            After:=targetWkb.Worksheets(targetWkb.Worksheets.Count))
        sh.Name = targetName
    Else
        For counter = sh.ListObjects.Count To 1 Step -1
            sh.ListObjects(counter).Delete
        Next
        sh.Cells.Clear
    End If

    Set FindOrCreateSheet = sh
End Function

'@label:format-export-sheet
'@sub-title Apply standard formatting to an exported sheet
'@details
'Sets font and alignment over the used range, autofits it, then sets the
'sheet visibility.
'
'The look of the sheet is cosmetic and a failure there is suppressed. The
'visibility write is not cosmetic and stands outside the suppression: Excel
'refuses to hide the last visible worksheet of a workbook and refuses a write
'to a protected workbook structure, and either one leaves the translation
'tables on show in a delivered linelist.
'@param sh Worksheet. The sheet to format.
'@param Hide Long. Sheet visibility constant.
Private Sub FormatExportSheet(ByVal sh As Worksheet, ByVal Hide As Long)
    Dim usedRng As Range

    Set usedRng = sh.UsedRange

    On Error Resume Next
    sh.Rows(1).Font.Bold = True
    usedRng.Font.Name = "Consolas"
    usedRng.HorizontalAlignment = xlHAlignLeft
    usedRng.VerticalAlignment = xlVAlignCenter
    usedRng.EntireRow.AutoFit
    usedRng.EntireColumn.AutoFit
    On Error GoTo 0

    sh.Visible = Hide
End Sub

'@label:check-requirements
'@sub-title Validate that the translation sheet has all required tables
'@details
'Ensures the worksheet is not Nothing and that all five translation
'ListObjects are present.
'@param translationSheet Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When the sheet is Nothing.
'@throws ProjectError.ElementNotFound When a required table is missing.
Private Sub CheckRequirements(ByVal translationSheet As Worksheet)
    Dim requiredTables As Variant
    Dim counter As Long
    Dim tableName As String
    Dim Lo As ListObject

    If translationSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Translation sheet is required."
    End If

    requiredTables = Array(TRADLLMSG, TRADLLSHAPES, TRADLLFORMS, TRADLLRIBBON, DICTTRANS)

    For counter = LBound(requiredTables) To UBound(requiredTables)
        tableName = requiredTables(counter)
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


'@section HiddenName writing
'===============================================================================
'@description Creating a hidden name and giving it its value.

'@label:ensure-and-set
'@sub-title Give a HiddenName its value, creating it when it is absent
'@details
'EnsureName writes the value only when it creates the name, so a name that is
'already there needs the SetValue. Creating one and setting it straight after
'costs a full RefersTo rewrite and a comment re-apply on every one of the 17
'names, which is why the two are branches here.
'@param store HiddenNames. The hidden names manager.
'@param nameId String. The hidden name identifier.
'@param value String. The value to set.
Private Sub EnsureAndSet(ByVal store As HiddenNames, ByVal nameId As String, _
                         ByVal value As String)
    If store.HasName(nameId) Then
        store.SetValue nameId, value
    Else
        store.EnsureName nameId, value, HiddenNameTypeString
    End If
End Sub

'@section Error Handling
'===============================================================================
'@description Centralised error raising.

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber ProjectError. The error constant.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
