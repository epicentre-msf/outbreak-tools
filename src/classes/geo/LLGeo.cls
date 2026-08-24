VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Geobase management class"

'@PredeclaredId
'@Folder("Geo")
'@ModuleDescription("Geobase management class")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing

Option Explicit

'@section Enumerations
'===============================================================================

'@details
'Geographic scope selector. GeoScopeAdmin picks the admin hierarchy,
'GeoScopeHF picks the health facility table, and GeoScopeBoth is used by
'ClearHistoric to name both historic tables at once.
Public Enum GeoScope
    GeoScopeAdmin = 0
    GeoScopeHF = 1
    GeoScopeBoth = 3
End Enum

'@details
'Admin-only scope selects one of the four administrative levels without
'including health facilities. Used by GeoLevel when geoType is
'GeoScopeAdmin.
Public Enum AdminOnlyScope
    LevelAdmin1 = 0
    LevelAdmin2 = 1
    LevelAdmin3 = 2
    LevelAdmin4 = 3
End Enum

'@class LLGeo
'@description
'Manages the geobase worksheet: the four-level administrative hierarchy, admin1
'through admin4, the health facility table, the level-name translations and the
'historic entries. GeoNames and GeoBaseName answer the names, GeoLevel walks
'the hierarchy through AutoFilter cascades, AdminCode answers the code of one
'entry, and HasNoData says whether the geobase carries anything. Translate
'rewrites the level labels, Clear and ClearHistoric empty the tables, Import
'reads a geobase in, and ExportToFile and ExportToWkb write it out as flat data
'or as the full structure. Update, UpdateDict and UpdateHistoric carry the
'translated labels across the linelist sheets and the dictionary, and
'SkippedSheets names the sheets an update left alone.
'
'The show/hide form takes its geo labels from the dictionary UpdateDict
'rewrites: ShowHide composes "<sub section>: <level label>" itself, every time
'it is built. A third routine here used to write that same string into
'per-sheet tables on the show/hide worksheet, which is how the FIRST show/hide
'design carried it. ShowHideStore deletes those tables on sight now, so the
'routine walked a sheet holding nothing it could read and raised 91 on every
'geobase import. It went on 2026-08-14 with its one caller.
'
'WHERE THE STATE OF THE GEOBASE LIVES
'-------------------------------------------------------------------------------
'Two hidden names on the geo worksheet hold two separate facts. RNG_GeoUpdated
'answers whether the geobase carries data and reads "empty" or "updated".
'RNG_GeoTranslated is a Boolean and answers whether the table headers carry
'translated labels. A workbook built before this split holds one combined
'string; the first read moves it to the two names and writes them back.
'
'WHERE THE FIVE LEVEL NAMES LIVE
'-------------------------------------------------------------------------------
'RNG_ADM1NAME through RNG_ADM4NAME and RNG_HFNAME are workbook-scoped hidden
'names. CrossTable writes "= RNG_ADM1NAME" as a worksheet formula on the
'spatial analysis sheet, and a workbook-scoped name resolves there. A geo sheet
'built before this change still carries them as cells; the factory reads each
'cell once, writes the hidden name and moves on. RNG_PastingGeoCol stays a real
'cell, because DesignerPreparation uses it as a paste anchor.
'
'WHO OWNS RNG_MetaLang
'-------------------------------------------------------------------------------
'This class reads RNG_MetaLang in ExportToWkb and writes it nowhere. The
'writers are InitTransfer, which stamps the meta language picked on the
'designer's Main sheet onto the designer geo sheet just before the export, and
'the geobase import round trip.
'@depends BetterArray, DropdownLists, HiddenNames, LLdictionary
'@depends Passwords

'@section State
'===============================================================================

Private Type TGeo
    geosh As Worksheet
    geoHidden As HiddenNames
    levelHidden As HiddenNames
    tableNames As BetterArray
    skipped As BetterArray
    adm1Label As String
    adm2Label As String
    adm3Label As String
    adm4Label As String
    hfLabel As String
    namesCacheLoaded As Boolean
    geoUpdated As String
    translated As Boolean
    flagsLoaded As Boolean
    lastFailure As String
    levelMemo As Collection
    sealed As Boolean
End Type

Private this As TGeo
Private Const CLASS_NAME As String = "LLGeo"

'Table name constants
Private Const ADMIN1 As String = "ADM1"
Private Const ADMIN2 As String = "ADM2"
Private Const ADMIN3 As String = "ADM3"
Private Const ADMIN4 As String = "ADM4"
Private Const HF As String = "HF"
Private Const HISTOGEO As String = "HISTOGEO"
Private Const HISTOHF As String = "HISTOHF"
Private Const NAMES As String = "NAMES"
Private Const METADATA As String = "METADATA"

'Hidden names on the geo worksheet
Private Const HN_GEO_UPDATED As String = "RNG_GeoUpdated"
Private Const HN_GEO_TRANSLATED As String = "RNG_GeoTranslated"
Private Const HN_GEO_NAME As String = "RNG_GeoName"
Private Const HN_META_LANG As String = "RNG_MetaLang"
Private Const HN_GEO_LANG_CODE As String = "RNG_GeoLangCode"

'The version of the designer, held as a workbook-scoped hidden name on the
'designer itself. DesignerPreparation.Prepare writes it, from the RNG_Version
'value name the owner keeps on the Dev worksheet. A designer that carries
'neither puts VERSION_NOT_FOUND on the metadata sheet.
Private Const HN_DESIGNER_VERSION As String = "RNG_DesignerVersion"
Private Const VERSION_NOT_FOUND As String = "(not found)"

'The two values RNG_GeoUpdated takes
Private Const FLAG_EMPTY As String = "empty"
Private Const FLAG_UPDATED As String = "updated"

'Workbook-scoped hidden names holding the five translated level labels
Private Const HN_ADM1_NAME As String = "RNG_ADM1NAME"
Private Const HN_ADM2_NAME As String = "RNG_ADM2NAME"
Private Const HN_ADM3_NAME As String = "RNG_ADM3NAME"
Private Const HN_ADM4_NAME As String = "RNG_ADM4NAME"
Private Const HN_HF_NAME As String = "RNG_HFNAME"

'The cell-based named range that stays a cell
Private Const RNG_PASTING_GEO_COL As String = "RNG_PastingGeoCol"

'Column name in T_NAMES that holds the level identifiers (adm1_name, etc.)
Private Const NAMES_LEVEL_COL As String = "level"

'Worksheet-level hidden name holding the kind of a linelist sheet
Private Const SHEET_TAG As String = "sheet_type"

'The key UniqueRangeValue dedupes with. The prefix is load-bearing twice over:
'a Collection refuses an empty key, and an empty cell is a real entry of a
'geobase column, so the prefix is what lets one be held. The length cap is
'Excel's own: a Collection key longer than 255 characters raises, and two place
'names that agree on their first 250 characters count as one here.
Private Const UNIQUE_KEY_PREFIX As String = "k:"
Private Const UNIQUE_KEY_MAXLENGTH As Long = 250

'How many answered levels the instance holds. The cascade asks for the same
'parents over and over -- a column filled with one admin1, a value re-picked, a
'typo corrected -- and each miss costs a filter of the whole geobase. The cap is
'what keeps a long session from holding a geobase twice over.
Private Const LEVEL_MEMO_MAX As Long = 16

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new LLGeo instance from a geobase worksheet.
'@details
'Factory method that validates all required tables, named ranges, and
'hidden names on the supplied worksheet, then constructs and returns a
'new LLGeo instance bound to that worksheet. This is the only supported
'way to create a usable LLGeo object because the class relies on
'specific ListObjects, named ranges, and HiddenNames being present.
'
'The two hidden name stores are built here, once each, and handed to the
'instance. Building one walks every Name of its container, and the class
'reads both stores on nearly every call.
'
'A LINELIST OLDER THAN THE HIDDEN NAME STORE
'-------------------------------------------------------------------------------
'A linelist generated before this version kept its geo metadata in plain named
'ranges, and the hidden name store is what replaced them. HiddenNames tracks a
'name only when its Visible property is False, so on such a file every one of
'the four names the checks demand reads as absent and the factory refuses a
'geobase that is entirely there. That is what the other-linelist export meets
'when a user points it at an old file.
'
'So the checks are optional. They stay on for everything that builds or edits a
'linelist, where a missing name is a real fault worth stopping on, and the
'export walks turn them off: an export reads the geobase tables and writes them
'out, and the tables are what it needs. Whatever the file cannot answer fails at
'the read that wants it, which names the missing piece instead of one error at
'the door.
'@param sh Worksheet. The geobase worksheet containing T_ADM1 through T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF, T_METADATA, and required named ranges/hidden names.
'@param runChecks Optional Boolean. When False the tables, named ranges and
'  hidden names are not checked before the instance is built. Defaults to True.
'  It is NOT named checkRequirements: CheckRequirements is a Sub of this class,
'  and a parameter of that name turns the call below into an index on a Boolean.
'@return LLGeo. A fully initialised LLGeo instance.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.ElementNotFound When a required table, hidden name or named range is missing and runChecks is True.
'@export
Public Function Create(ByVal sh As Worksheet, _
                       Optional ByVal runChecks As Boolean = True) As LLGeo
    Dim instance As LLGeo
    Dim geoStore As HiddenNames
    Dim levelStore As HiddenNames

    If sh Is Nothing Then ThrowError ObjectNotInitialized, "Geo sheet not defined"

    Set geoStore = HiddenNames.Create(sh)
    Set levelStore = HiddenNames.Create(sh.Parent)
    If runChecks Then CheckRequirements sh, geoStore, levelStore

    Set instance = New LLGeo
    Set instance.Wksh = sh
    Set instance.GeoStore = geoStore
    Set instance.LevelStore = levelStore
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes.
'@details
'Marks the instance as sealed so setup-only setters raise if called again.
'Called by the factory after the last setup write.
Public Sub Seal()
    this.sealed = True
End Sub

'@sub-title Guard a setup-only setter against post-seal writes.
'@param propName String. The property name used in the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:wksh
'@prop-title Get the geobase worksheet.
'@details
'Returns the worksheet that holds all geobase ListObjects and named ranges.
'This worksheet is set during factory creation and holds for the lifetime of
'the object.
'@return Worksheet. The geobase worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.geosh
End Property

'@label:wksh-set
'@prop-title Set the geobase worksheet.
'@details
'Assigns the internal worksheet reference. Called from the factory Create
'method after validation, and only from there.
'@param sh Worksheet. The geobase worksheet to bind to this instance.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.geosh = sh
End Property

'@label:geo-store-set
'@prop-title Set the worksheet hidden name store.
'@details
'Setup-only setter. The factory builds the store while it validates the
'worksheet and hands the same instance over, so the Names collection is
'walked once per Create.
'@param store HiddenNames. The store over the geobase worksheet.
Public Property Set GeoStore(ByVal store As HiddenNames)
    GuardNotSealed "GeoStore"
    Set this.geoHidden = store
End Property

'@label:level-store-set
'@prop-title Set the workbook hidden name store.
'@details
'Setup-only setter. Holds the store over the workbook that owns the geobase
'worksheet, which is where the five translated level labels live.
'@param store HiddenNames. The store over the workbook.
Public Property Set LevelStore(ByVal store As HiddenNames)
    GuardNotSealed "LevelStore"
    Set this.levelHidden = store
End Property

'@section Hidden Names Stores
'===============================================================================

'@label:geo-hidden-store
'@sub-title HiddenNames manager for the geo worksheet.
'@details
'Returns the store the factory built. Builds one on demand when the instance
'was made another way. It holds RNG_GeoUpdated, RNG_GeoTranslated,
'RNG_GeoName, RNG_MetaLang and RNG_GeoLangCode.
'@return HiddenNames. The worksheet-level hidden names manager.
Private Function GeoHiddenStore() As HiddenNames
    If this.geoHidden Is Nothing Then
        Set this.geoHidden = HiddenNames.Create(Wksh())
    End If
    Set GeoHiddenStore = this.geoHidden
End Function

'@label:level-name-store
'@sub-title HiddenNames manager for the workbook holding the geo worksheet.
'@details
'Returns the store the factory built. The five level labels are
'workbook-scoped because CrossTable resolves them from a worksheet formula
'on the spatial analysis sheet.
'@return HiddenNames. The workbook-level hidden names manager.
Private Function LevelNameStore() As HiddenNames
    If this.levelHidden Is Nothing Then
        Set this.levelHidden = HiddenNames.Create(Wksh().Parent)
    End If
    Set LevelNameStore = this.levelHidden
End Function

'@section Geobase Flags
'===============================================================================

'@label:ensure-flags-cache
'@sub-title Read the two geobase flags once per instance.
'@details
'Reads RNG_GeoUpdated and RNG_GeoTranslated into the instance. A geo sheet
'written before the two facts were split carries one combined string in
'RNG_GeoUpdated, such as "updated, translated". That string is parsed once
'here, the two clean values are written back, and no reader parses a string
'again.
Private Sub EnsureFlagsCache()
    Dim store As HiddenNames
    Dim combined As String

    If this.flagsLoaded Then Exit Sub

    Set store = GeoHiddenStore()

    If store.HasName(HN_GEO_TRANSLATED) Then
        this.geoUpdated = store.ValueAsString(HN_GEO_UPDATED)
        this.translated = store.ValueAsBoolean(HN_GEO_TRANSLATED)
    Else
        combined = store.ValueAsString(HN_GEO_UPDATED)
        this.translated = (InStr(1, combined, ", translated") > 0)
        combined = Replace(combined, ", not translated", vbNullString)
        this.geoUpdated = Replace(combined, ", translated", vbNullString)

        store.EnsureName HN_GEO_TRANSLATED, this.translated, HiddenNameTypeBoolean
        store.SetValue HN_GEO_TRANSLATED, this.translated
        store.SetValue HN_GEO_UPDATED, this.geoUpdated
    End If

    this.flagsLoaded = True
End Sub

'@label:reset-flags-cache
'@sub-title Drop the cached flags so the next read goes to the store.
'@details
'Called after anything that writes the hidden names behind the class's back,
'which today is the name import Import runs off a geobase Metadata sheet.
Private Sub ResetFlagsCache()
    this.flagsLoaded = False
End Sub

'@label:geo-updated-flag
'@sub-title The data flag of the geobase.
'@return String. "empty" or "updated".
Private Function GeoUpdatedFlag() As String
    EnsureFlagsCache
    GeoUpdatedFlag = this.geoUpdated
End Function

'@label:set-geo-updated-flag
'@sub-title Write the data flag and keep the cache in step.
'@param flagValue String. "empty" or "updated".
Private Sub SetGeoUpdatedFlag(ByVal flagValue As String)
    EnsureFlagsCache
    GeoHiddenStore().SetValue HN_GEO_UPDATED, flagValue
    this.geoUpdated = flagValue
End Sub

'@label:is-translated
'@sub-title Whether the table headers carry translated labels.
'@return Boolean. True when the headers are translated.
Private Function IsTranslated() As Boolean
    EnsureFlagsCache
    IsTranslated = this.translated
End Function

'@label:set-translated-flag
'@sub-title Write the translation flag and keep the cache in step.
'@param flagValue Boolean. True when the headers carry translated labels.
Private Sub SetTranslatedFlag(ByVal flagValue As Boolean)
    Dim store As HiddenNames

    EnsureFlagsCache
    Set store = GeoHiddenStore()
    store.EnsureName HN_GEO_TRANSLATED, flagValue, HiddenNameTypeBoolean
    store.SetValue HN_GEO_TRANSLATED, flagValue
    this.translated = flagValue
End Sub

'@section Table Names
'===============================================================================

'@label:table-names
'@prop-title Ordered list of all geobase table names.
'@details
'Returns the eight canonical ListObject names of the geobase worksheet:
'T_ADM1, T_ADM2, T_ADM3, T_ADM4, T_HF, T_NAMES, T_HISTOGEO and T_HISTOHF.
'The order matches the admin hierarchy followed by support tables, and the
'lower bound is 0, which is what GeoColumnIndex and GeoLevel index against.
'The array is built once per instance and handed back as it stands, because
'GeoColumnIndex asks for it on every column resolution.
'@return BetterArray. Ordered list of table name strings.
Private Property Get TableNames() As BetterArray
    If this.tableNames Is Nothing Then
        Set this.tableNames = New BetterArray
        this.tableNames.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
                             "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF
    End If
    Set TableNames = this.tableNames
End Property

'@label:table-name-at
'@sub-title One geobase table name by position.
'@param position Long. Zero-based position in the table name list.
'@return String. The ListObject name at that position.
Private Function TableNameAt(ByVal position As Long) As String
    TableNameAt = CStr(TableNames().Item(position))
End Function

'@section Level Names
'===============================================================================

'@label:level-raw-name
'@sub-title The raw column name of one level.
'@param position Long. 1 to 4 for the admin levels, 5 for health facility.
'@return String. "adm1_name" to "adm4_name", or "hf_name".
Private Function LevelRawName(ByVal position As Long) As String
    If position = 5 Then
        LevelRawName = "hf_name"
    Else
        LevelRawName = "adm" & position & "_name"
    End If
End Function

'@label:level-name-id
'@sub-title The hidden name holding the label of one level.
'@param rawName String. The raw column name of the level.
'@return String. The hidden name identifier, or an empty string for an unknown level.
Private Function LevelNameId(ByVal rawName As String) As String
    Select Case rawName
    Case "adm1_name"
        LevelNameId = HN_ADM1_NAME
    Case "adm2_name"
        LevelNameId = HN_ADM2_NAME
    Case "adm3_name"
        LevelNameId = HN_ADM3_NAME
    Case "adm4_name"
        LevelNameId = HN_ADM4_NAME
    Case "hf_name"
        LevelNameId = HN_HF_NAME
    End Select
End Function

'@label:ensure-names-cache
'@sub-title Ensure the five level labels are held on the instance.
'@details
'Reads the five workbook-scoped hidden names into five String fields. A
'level with no stored label answers its own raw column name, which is what
'every reader falls back to.
Private Sub EnsureNamesCache()
    If this.namesCacheLoaded Then Exit Sub

    this.adm1Label = ReadLevelName("adm1_name")
    this.adm2Label = ReadLevelName("adm2_name")
    this.adm3Label = ReadLevelName("adm3_name")
    this.adm4Label = ReadLevelName("adm4_name")
    this.hfLabel = ReadLevelName("hf_name")

    this.namesCacheLoaded = True
End Sub

'@label:read-level-name
'@sub-title Read one level label from the workbook store.
'@param rawName String. The raw column name of the level.
'@return String. The stored label, or the raw name when nothing is stored.
Private Function ReadLevelName(ByVal rawName As String) As String
    Dim stored As String

    stored = LevelNameStore().ValueAsString(LevelNameId(rawName))
    If LenB(stored) = 0 Then stored = rawName
    ReadLevelName = stored
End Function

'@label:write-level-name
'@sub-title Write one level label into the workbook store.
'@param rawName String. The raw column name of the level.
'@param label String. The translated label to store.
Private Sub WriteLevelName(ByVal rawName As String, ByVal label As String)
    Dim nameId As String
    Dim store As HiddenNames

    nameId = LevelNameId(rawName)
    If LenB(nameId) = 0 Then Exit Sub

    Set store = LevelNameStore()
    store.EnsureName nameId, label, HiddenNameTypeString
    store.SetValue nameId, label
End Sub

'@label:reset-names-cache
'@sub-title Reset the level label cache.
'@details
'Marks the cache as unloaded so the next GeoNames access rebuilds it from the
'store. Called after anything that changes translation state.
Private Sub ResetNamesCache()
    this.namesCacheLoaded = False
End Sub

'@label:update-level-names
'@sub-title Refresh the five level labels from T_NAMES.
'@details
'Reads the current language code from RNG_GeoLangCode, finds the matching
'column in T_NAMES, and writes each level label into its workbook-scoped
'hidden name. The whole T_NAMES body is read in one crossing and all five
'levels are resolved in one pass over it. Exits when no language code is set
'or when T_NAMES lacks the level column or the language column.
Private Sub UpdateLevelNames()
    Dim sh As Worksheet
    Dim langCode As String
    Dim loNames As ListObject
    Dim levelCol As Long
    Dim langCol As Long
    Dim block As Variant
    Dim rowIdx As Long
    Dim rawName As String
    Dim label As String

    langCode = GeoHiddenStore().ValueAsString(HN_GEO_LANG_CODE)
    If LenB(langCode) = 0 Then Exit Sub

    Set sh = Wksh()
    Set loNames = sh.ListObjects("T_" & NAMES)
    If loNames.ListRows.Count = 0 Then Exit Sub

    levelCol = ColumnIndex(NAMES_LEVEL_COL, loNames.HeaderRowRange)
    If levelCol < 0 Then Exit Sub

    langCol = ColumnIndex(langCode, loNames.HeaderRowRange)
    If langCol < 0 Then Exit Sub

    block = RangeValues(loNames.DataBodyRange)

    For rowIdx = LBound(block, 1) To UBound(block, 1)
        rawName = CellText(block(rowIdx, levelCol))

        Select Case rawName
        Case "adm1_name", "adm2_name", "adm3_name", "adm4_name", "hf_name"
            label = CellText(block(rowIdx, langCol))
            If LenB(label) > 0 Then WriteLevelName rawName, label
        End Select
    Next
End Sub

'@label:geo-names
'@prop-title Translated display name for a raw admin column name.
'@details
'Converts a raw admin column name (such as "adm1_name" or "hf_name") to its
'stored display label. An unknown key comes back unchanged, and so does a
'level whose label has never been written.
'@param admname String. Raw column name to translate.
'@return String. Display label, or the raw name when no label is stored.
'@export
Public Property Get GeoNames(ByVal admname As String) As String
    EnsureNamesCache

    Select Case admname
    Case "adm1_name"
        GeoNames = this.adm1Label
    Case "adm2_name"
        GeoNames = this.adm2Label
    Case "adm3_name"
        GeoNames = this.adm3Label
    Case "adm4_name"
        GeoNames = this.adm4Label
    Case "hf_name"
        GeoNames = this.hfLabel
    Case Else
        GeoNames = admname
    End Select
End Property

'@label:geo-base-name
'@prop-title The name of the geobase the worksheet was filled from.
'@details
'Import stores the source file name in RNG_GeoName. LLExporter reads this
'when it builds the file name of an exported geobase.
'@return String. The stored geobase name, or an empty string.
'@export
Public Property Get GeoBaseName() As String
    GeoBaseName = GeoHiddenStore().ValueAsString(HN_GEO_NAME)
End Property

'@section Column Helpers
'===============================================================================

'@label:column-index
'@sub-title Find the 1-based column index of a name within a header range.
'@details
'Searches a header row range for an exact case-sensitive match of the given
'column name. When inRange is True the returned index is relative to the range
'start (1-based); when False the absolute worksheet column number is returned.
'Returns -1 when no match is found.
'
'Every option of the search is pinned. Range.Find inherits LookIn, SearchOrder
'and SearchDirection from the last search of the Excel session, including one
'the user ran from the Find dialog, and a prior LookIn:=xlComments makes every
'column lookup in this class miss.
'@param colName String. The column name to search for.
'@param headerRng Range. The header row range to search in.
'@param inRange Optional Boolean. True to return relative position, False for absolute column number. Defaults to True.
'@return Long. Column index (1-based), or -1 if not found.
Private Function ColumnIndex(ByVal colName As String, ByVal headerRng As Range, _
                             Optional ByVal inRange As Boolean = True) As Long
    Dim matchCell As Range

    Set matchCell = headerRng.Find(What:=colName, _
                                   lookAt:=xlWhole, _
                                   LookIn:=xlValues, _
                                   SearchOrder:=xlByColumns, _
                                   SearchDirection:=xlNext, _
                                   MatchCase:=True)

    If matchCell Is Nothing Then
        ColumnIndex = -1
    ElseIf inRange Then
        ColumnIndex = matchCell.Column - headerRng.Column + 1
    Else
        ColumnIndex = matchCell.Column
    End If
End Function

'@label:geo-column-index
'@sub-title Resolve the column index for a geo admin table.
'@details
'Looks up a column name in the header of the specified admin-level ListObject.
'When the headers carry translated labels the column name is converted first.
'Raises when the column cannot be found in the table header.
'@param colName String. The raw column name to look up.
'@param adminLevel Byte. Admin level index selecting which table to search (0-3 for T_ADM1 through T_ADM4).
'@return Long. 1-based column index within the table.
'@throws ProjectError.ElementNotFound When the column is not found in the admin table.
Private Function GeoColumnIndex(ByVal colName As String, ByVal adminLevel As Byte) As Long
    Dim sh As Worksheet
    Dim loName As String
    Dim actualColName As String
    Dim colIdx As Long

    Set sh = Wksh()
    loName = TableNameAt(adminLevel)
    actualColName = colName

    If IsTranslated() Then actualColName = GeoNames(colName)

    colIdx = ColumnIndex(actualColName, sh.ListObjects(loName).HeaderRowRange)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & actualColName & " not found in table " & loName

    GeoColumnIndex = colIdx
End Function

'@label:hf-column-index
'@sub-title Resolve the column index for the HF table.
'@details
'Looks up a column name in the header of the T_HF ListObject. When the headers
'carry translated labels the column name is converted first. Raises when the
'column cannot be found.
'@param colName String. The raw column name to look up.
'@return Long. 1-based column index within the T_HF table.
'@throws ProjectError.ElementNotFound When the column is not found in the HF table.
Private Function HFColumnIndex(ByVal colName As String) As Long
    Dim colIdx As Long

    colIdx = HFColumnIndexOrMissing(colName)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & colName & " not found in table T_HF"

    HFColumnIndex = colIdx
End Function

'@label:hf-column-index-or-missing
'@sub-title Resolve a column of T_HF, answering -1 when it has none.
'@details
'The health facility table of a production geobase carries four levels and
'the one of a test fixture can carry fewer, so the header translation walks
'the levels and skips the ones the table has no column for.
'@param colName String. The raw column name to look up.
'@return Long. 1-based column index within T_HF, or -1 when absent.
Private Function HFColumnIndexOrMissing(ByVal colName As String) As Long
    Dim actualColName As String

    actualColName = colName
    If IsTranslated() Then actualColName = GeoNames(colName)

    HFColumnIndexOrMissing = ColumnIndex(actualColName, _
                                         Wksh().ListObjects("T_" & HF).HeaderRowRange)
End Function

'@label:cell-text
'@sub-title Read one variant as text, answering an empty string for an error value.
'@details
'CStr on a cell holding #N/A raises 13. Every read of a cell value in this
'class goes through here.
'@param cellValue Variant. The value read from a cell.
'@return String. The value as text, or an empty string when it is an error.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:range-values
'@sub-title Read a range into a two-dimensional Variant array in one crossing.
'@details
'Range.Value answers a scalar for a single cell and a 2-D array for anything
'larger, so a single cell is wrapped into a 1x1 array here and every caller
'reads one shape.
'@param rng Range. The range to read.
'@return Variant. A two-dimensional array of the values.
Private Function RangeValues(ByVal rng As Range) As Variant
    Dim block As Variant

    If rng.Cells.Count = 1 Then
        ReDim block(1 To 1, 1 To 1)
        block(1, 1) = rng.Value
    Else
        block = rng.Value
    End If

    RangeValues = block
End Function

'@label:unique-range-value
'@sub-title Extract unique values from a range into a BetterArray.
'@details
'Reads one Variant array per area of the range and dedupes in memory against a
'keyed Collection, then pushes the values in encounter order. The range handed
'in is the visible part of a filtered column, which on a large geobase is tens
'of thousands of cells.
'
'WHY A COLLECTION AND NOT A STRING INDEX
'-------------------------------------------------------------------------------
'The index used to be one long string with every value already seen in it, and
'every cell was an InStr over the whole of it. That is quadratic in the number
'of distinct values, and the string was rebuilt on every hit, so a level with a
'few thousand children cost far more than the filter that produced it. A
'Collection key is hashed, so the test is flat however many values arrive.
'Dictionary is missing on Mac Excel; a keyed Collection is the replacement this
'project uses.
'
'The dedupe is now CASE INSENSITIVE, because Collection keys are. That follows
'the AutoFilter above it, which matched these rows case insensitively in the
'first place, so two spellings of one place name were already being treated as
'one by the time they got here.
'@param rng Range. The range of cells to extract unique values from.
'@return BetterArray. Deduplicated values in encounter order.
Private Function UniqueRangeValue(ByVal rng As Range) As BetterArray
    Dim retVal As BetterArray
    Dim areaRng As Range
    Dim block As Variant
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim cellValue As Variant
    Dim entry As String
    Dim seen As Collection

    Set retVal = New BetterArray
    Set seen = New Collection

    For Each areaRng In rng.Areas
        block = RangeValues(areaRng)

        'The suppression covers the keyed Add alone, and the raise IS the test:
        'a Collection refuses a key it already holds. It is armed once per area
        'rather than once per cell, because a geobase column crosses tens of
        'thousands of cells and each arming costs.
        On Error Resume Next
        For rowIdx = LBound(block, 1) To UBound(block, 1)
            For colIdx = LBound(block, 2) To UBound(block, 2)
                cellValue = block(rowIdx, colIdx)
                entry = CellText(cellValue)

                'Cleared first, so a raise from the previous Push cannot read as
                'a duplicate key here and drop a value that belongs in the list.
                Err.Clear
                seen.Add entry, _
                         UNIQUE_KEY_PREFIX & Left$(entry, UNIQUE_KEY_MAXLENGTH)
                If Err.Number = 0 Then retVal.Push cellValue
            Next
        Next
        On Error GoTo 0
    Next

    Set UniqueRangeValue = retVal
End Function

'@section Level memo
'===============================================================================
'One answered level per slot, so the cascade pays the AutoFilter once per set of
'parents rather than once per keystroke. Import and Clear empty it, and the
'whole instance is dropped by EventLinelist.ResetCaches after a geobase import,
'so nothing here can outlive the geobase it was read from.

'@label:level-memo-key
'@sub-title The identity of one GeoLevel answer.
'@details
'Scope and level, then the parents that narrow it. The parents arrive as three
'shapes -- nothing for admin 1, a scalar for admin 2, a BetterArray below that
'-- and each becomes text here. A parent carrying the separator itself would
'read as two parents, so the separator is a character no place name holds.
'@param adminLevel Byte. Admin level index.
'@param geoType Byte. GeoScopeAdmin or GeoScopeHF.
'@param adminNames Variant. The parent selections as GeoLevel received them.
'@return String. The key of that answer.
Private Function LevelMemoKey(ByVal adminLevel As Byte, ByVal geoType As Byte, _
                              ByRef adminNames As Variant) As String
    Const KEY_SEPARATOR As String = vbNullChar

    Dim key As String
    Dim counter As Long

    key = CStr(geoType) & KEY_SEPARATOR & CStr(adminLevel)

    If IsObject(adminNames) Then
        If TypeName(adminNames) = "BetterArray" Then
            For counter = 1 To adminNames.Length
                key = key & KEY_SEPARATOR & CStr(adminNames.Item(counter))
            Next
        End If
    Else
        key = key & KEY_SEPARATOR & CStr(adminNames)
    End If

    LevelMemoKey = key
End Function

'@label:level-memo-read
'@sub-title The held answer for one key, as a list of its own.
'@details
'A fresh BetterArray every time, filled from the items held under the key. Two
'reasons it is not the held object itself: VarWriter clears and refills what
'GeoLevel hands it while it writes the geo1 column, so one caller would empty
'the memo for every other; and a caller must not be able to tell a hit from a
'miss. Setting Items on a new BetterArray adopts the bounds of what is stored,
'so a hit answers the same shape the filter would have built.
'@param key String. The key built by LevelMemoKey.
'@return BetterArray. The held answer, or Nothing when the key is not held.
Private Function LevelMemoRead(ByVal key As String) As BetterArray
    Dim storedItems As Variant
    Dim answer As BetterArray

    If this.levelMemo Is Nothing Then Exit Function

    'A Collection raises 5 on a key it does not hold, and that raise is the
    'test. The suppression covers the read alone.
    On Error Resume Next
    storedItems = this.levelMemo.Item(key)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set answer = New BetterArray
    answer.Items = storedItems
    Set LevelMemoRead = answer
End Function

'@label:level-memo-write
'@sub-title Hold one answer under its key.
'@details
'The items are held rather than the BetterArray, for the reason LevelMemoRead
'gives. An empty answer is not held: an empty list has no one array shape to
'store, and a level with no children at all is a geobase fault rather than
'something a user meets on the way down a column.
'
'A full memo is emptied whole. Dropping the oldest entry would need a second
'structure to carry the order, and the cap is a guard against a session that
'never ends rather than a working limit.
'@param key String. The key built by LevelMemoKey.
'@param values BetterArray. The answer to hold.
Private Sub LevelMemoWrite(ByVal key As String, ByVal values As BetterArray)
    If values Is Nothing Then Exit Sub
    If values.Length = 0 Then Exit Sub

    If this.levelMemo Is Nothing Then ResetLevelMemo
    If this.levelMemo.Count >= LEVEL_MEMO_MAX Then ResetLevelMemo

    'The key can only be there already when two calls raced the read above,
    'which one thread cannot do; the suppression is what keeps a surprise from
    'reaching the caller as a geobase failure.
    On Error Resume Next
    this.levelMemo.Add values.Items, key
    Err.Clear
    On Error GoTo 0
End Sub

'@label:reset-level-memo
'@sub-title Drop every held level answer.
Private Sub ResetLevelMemo()
    Set this.levelMemo = New Collection
End Sub

'@label:hf-range
'@sub-title Get the data body range for a named column in the HF table.
'@details
'Resolves the column index for the given column name in T_HF and returns the
'DataBodyRange of that ListColumn. Used by GeoLevel when retrieving health
'facility data at admin level 1.
'@param colName String. The raw column name whose data body range is needed.
'@return Range. The DataBodyRange of the matching column in T_HF.
Private Function HFRange(ByVal colName As String) As Range
    Dim colIdx As Long
    colIdx = HFColumnIndex(colName)
    Set HFRange = Wksh().ListObjects("T_" & HF).ListColumns(colIdx).DataBodyRange()
End Function

'@label:clear-list-filter
'@sub-title Drop the AutoFilter of a table when it carries one.
'@param Lo ListObject. The table to clear.
Private Sub ClearListFilter(ByVal Lo As ListObject)
    If Lo Is Nothing Then Exit Sub
    If Lo.AutoFilter Is Nothing Then Exit Sub
    If Not Lo.AutoFilter.FilterMode Then Exit Sub
    Lo.AutoFilter.ShowAllData
End Sub

'@section GeoLevel
'===============================================================================

'@label:guard-level-names
'@sub-title Validate the parent selections handed to GeoLevel.
'@details
'Admin 1 needs no parent. Admin 2 needs the name of one admin 1 and takes any
'scalar, because two call sites in EventLinelist hand over a raw cell value:
'an empty cell reads as Empty and a geobase whose admin 1 labels are numeric
'codes reads as Double. Admin 3 and admin 4 need a BetterArray of the parent
'names, one entry per level above them, with 1 as its lower bound.
'@param adminLevel Byte. Admin level index (0-3 for admin1-admin4).
'@param adminNames Variant. The parent selections handed in by the caller.
'@throws ProjectError.InvalidArgument When the parent selections are the wrong shape.
Private Sub GuardLevelNames(ByVal adminLevel As Byte, ByRef adminNames As Variant)
    If adminLevel = LevelAdmin1 Then Exit Sub

    If adminLevel = LevelAdmin2 Then
        If IsObject(adminNames) Then _
            ThrowError InvalidArgument, _
                "Admin 2 needs the name of one admin 1 as a single value"
        Exit Sub
    End If

    If Not TypeName(adminNames) = "BetterArray" Then _
        ThrowError InvalidArgument, _
            "Admin " & (CLng(adminLevel) + 1) & " needs a table of the " & _
            CLng(adminLevel) & " admin names above it"

    If adminNames.Length <> CLng(adminLevel) Then _
        ThrowError InvalidArgument, _
            "You should provide a table of length " & CLng(adminLevel) & _
            " for admin " & (CLng(adminLevel) + 1)

    If adminNames.LowerBound <> 1 Then _
        ThrowError InvalidArgument, _
            "Please ensure your BetterArray has 1 as LowerBound, it is mandatory"
End Sub

'@label:geo-level
'@prop-title List of values at a given admin level.
'@details
'Returns a BetterArray of unique geographic names at the requested admin level,
'optionally filtered by parent admin selections. For admin1 the full unfiltered
'list is returned. For admin2 through admin4 the caller supplies parent
'selections so the result cascades hierarchically via AutoFilter. The geoType
'parameter switches between the admin hierarchy and the health facility table.
'Calculation is turned off while the filtering runs.
'
'An empty geobase answers an empty list, and it answers before any handler is
'armed. Every other failure is raised to the caller, and the AutoFilter of the
'table is dropped on the way out whichever way the call ends.
'@param adminLevel Byte. Admin level index (0-3 for admin1-admin4).
'@param geoType Optional Byte. GeoScope value: GeoScopeAdmin (0) or GeoScopeHF (1). Defaults to GeoScopeAdmin.
'@param adminNames Optional Variant. Parent admin selections: a single value for level 2, a BetterArray for levels 3-4. Defaults to vbNullString.
'@return BetterArray. Unique values at the requested level, or an empty array when the geobase is empty.
'@throws ProjectError.InvalidArgument When adminNames has the wrong shape or length.
'@throws ProjectError.ErrorUnexpectedState When the ListObject has no data.
'@throws ProjectError.ElementNotFound When a geo column is missing from the table.
'@export
Public Property Get GeoLevel(ByVal adminLevel As Byte, _
                              Optional ByVal geoType As Byte = GeoScopeAdmin, _
                              Optional adminNames As Variant = vbNullString) As BetterArray

    Dim sh As Worksheet
    Dim outLevels As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim rng As Range
    Dim loName As String
    Dim colIdx As Long
    Dim counter As Long
    Dim returnedRng As Range
    Dim prevCalc As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String
    Dim memoKey As String
    Dim memoAnswer As BetterArray

    Set sh = Wksh()
    Set outLevels = New BetterArray

    'An empty geobase answers an empty list, and it answers it here, outside
    'the handler below.
    If HasNoData() Then
        Set GeoLevel = outLevels
        Exit Property
    End If

    If geoType = GeoScopeAdmin Then
        loName = TableNameAt(adminLevel)
    Else
        loName = TableNameAt(4) 'T_HF
    End If

    GuardLevelNames adminLevel, adminNames

    'The held answer, if this level and these parents were asked for already.
    'It sits below the guard, so a caller handing the wrong shape still hears
    'about it rather than being answered from memory.
    memoKey = LevelMemoKey(adminLevel, geoType, adminNames)
    Set memoAnswer = LevelMemoRead(memoKey)
    If Not memoAnswer Is Nothing Then
        Set GeoLevel = memoAnswer
        Exit Property
    End If

    prevCalc = sh.EnableCalculation
    sh.EnableCalculation = False

    On Error GoTo CleanFail

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range

    If Lo.ListRows.Count = 0 Then
        ThrowError ErrorUnexpectedState, "The listobject " & loName & " has no data"
    End If

    Select Case geoType

    Case GeoScopeAdmin

        Select Case adminLevel

        Case LevelAdmin1
            colIdx = GeoColumnIndex("adm1_name", LevelAdmin1)
            Set rng = Lo.ListColumns(colIdx).DataBodyRange
            outLevels.FromExcelRange rng, DetectLastRow:=False, DetectLastColumn:=False

        Case LevelAdmin2
            colIdx = GeoColumnIndex("adm1_name", LevelAdmin2)
            LoRng.AutoFilter Field:=colIdx, Criteria1:=CStr(adminNames)
            colIdx = GeoColumnIndex("adm2_name", LevelAdmin2)

        Case LevelAdmin3, LevelAdmin4
            For counter = 1 To adminNames.Length
                colIdx = GeoColumnIndex("adm" & counter & "_name", adminLevel)
                LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames.Item(counter)
            Next
            colIdx = GeoColumnIndex("adm" & (CLng(adminLevel) + 1) & "_name", adminLevel)
        End Select

    Case GeoScopeHF

        Select Case adminLevel

        Case LevelAdmin1
            Set rng = HFRange("adm1_name")
            Set outLevels = UniqueRangeValue(rng)

        Case LevelAdmin2
            colIdx = HFColumnIndex("adm1_name")
            LoRng.AutoFilter Field:=colIdx, Criteria1:=CStr(adminNames)
            colIdx = HFColumnIndex("adm2_name")

        Case LevelAdmin3, LevelAdmin4
            For counter = 1 To adminNames.Length
                colIdx = HFColumnIndex("adm" & counter & "_name")
                LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames.Item(counter)
            Next
            If adminLevel = LevelAdmin3 Then
                colIdx = HFColumnIndex("adm3_name")
            Else
                colIdx = HFColumnIndex("hf_name")
            End If
        End Select
    End Select

    'For levels beyond admin1, read the filtered cells and then drop the filter
    If adminLevel <> LevelAdmin1 Then
        Set returnedRng = Nothing

        On Error Resume Next
        Set returnedRng = Lo.ListColumns(colIdx).DataBodyRange _
                            .SpecialCells(xlCellTypeVisible)
        On Error GoTo CleanFail

        If Not returnedRng Is Nothing Then Set outLevels = UniqueRangeValue(returnedRng)
        ClearListFilter Lo
    End If

    sh.EnableCalculation = prevCalc
    LevelMemoWrite memoKey, outLevels
    Set GeoLevel = outLevels
    Exit Property

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    Resume RaiseAfterFail

RaiseAfterFail:
    On Error Resume Next
    ClearListFilter Lo
    sh.EnableCalculation = prevCalc
    On Error GoTo 0
    Err.Raise errNumber, errSource, errMessage
End Property

'@section Translate
'===============================================================================

'@label:translate
'@sub-title Translate or revert geobase table headers.
'@details
'Applies or reverts translated display labels on all geobase table headers
'(admin1 through admin4 and health facility). When rawNames is False, headers
'take the labels stored for the five levels. When True, headers go back to the
'raw column names. The method exits when the headers already carry what is
'asked for, and it refreshes the level labels from T_NAMES only once it knows
'it has work to do.
'
'The admin table headers are written one array per table. The health facility
'headers are resolved by name, because that table carries its levels in an
'order the class does not fix.
'@param rawNames Optional Boolean. True to revert to raw names, False to apply the stored labels. Defaults to False.
'@export
'@label:try-translate
'@sub-title Translate the headers, and carry on when the sheet will not take it
'@details
'Translate writes through HiddenNames and into the header rows of the admin
'tables. Two things stop it on a real file: a linelist older than the hidden
'name store, whose geo names are visible and therefore invisible to the store,
'and a source workbook opened read-only, which an export walk always does.
'
'Neither is a reason to lose a geobase. The refusal is recorded in words and
'written to the linelist''s own log, and the caller carries on.
'THE READ OF THE FLAG IS INSIDE THE GUARD, AND HAS TO BE
'-------------------------------------------------------------------------------
'IsTranslated looks like a read and is not. EnsureFlagsCache, underneath it,
'writes the flag back when the geo sheet does not already carry it -- EnsureName
'then two SetValue calls, one of them onto RNG_GeoUpdated. On a linelist older
'than the hidden name store that name is present and VISIBLE, the store refuses
'to see it, and the write raises. So the refusal used to land one line above the
'guard and escape it. Both the read and the translate sit inside now.
'@param rawNames Boolean. True to put the raw names back, False for the labels.
'@param wasTranslated Boolean. Receives the state the headers were in, and False
'  whenever it could not be established -- which is what the restore paths want,
'  since nothing was changed and nothing needs putting back.
'@return Boolean. True when the headers were changed, False when they were left.
Private Function TryTranslate(ByVal rawNames As Boolean, _
                              ByRef wasTranslated As Boolean) As Boolean

    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    On Error GoTo TranslateRefused
    wasTranslated = IsTranslated()
    Translate rawNames:=rawNames
    TryTranslate = True
    Exit Function

TranslateRefused:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    wasTranslated = False
    ReportTranslateRefusal rawNames, errNumber, errSource, errMessage
End Function

'@label:report-translate-refusal
'@sub-title Put in words, in lastFailure, that the headers were left as they were
'@details
'This class does not reach a log, and it must not start: it is hand-exported
'into the designer, so a dependency added here has to be carried by hand into
'another workbook. lastFailure is how it already talks to whoever called it, and
'the caller is the one holding a log.
'
'LLExporter.ExportGeo reads this after a geobase export that WORKED, and writes
'the warning to the linelist''s own log from there.
'@param rawNames Boolean. Which direction was refused.
'@param errNumber Long. The error number.
'@param errSource String. The error source.
'@param errMessage String. The error description.
Private Sub ReportTranslateRefusal(ByVal rawNames As Boolean, _
                                   ByVal errNumber As Long, _
                                   ByVal errSource As String, _
                                   ByVal errMessage As String)

    this.lastFailure = "The geobase headers were left as they were: putting the " & _
                       IIf(rawNames, "raw names", "translated labels") & _
                       " on them raised error " & CStr(errNumber) & " from " & _
                       errSource & ": " & errMessage & ". The exported file " & _
                       "carries the data and the headers it already had."
End Sub

Public Sub Translate(Optional ByVal rawNames As Boolean = False)
    Dim sh As Worksheet
    Dim counter As Long
    Dim counter2 As Long
    Dim Lo As ListObject
    Dim hRng As Range
    Dim headerValues As Variant
    Dim rawName As String

    If (Not rawNames) And IsTranslated() Then Exit Sub
    If rawNames And (Not IsTranslated()) Then Exit Sub

    UpdateLevelNames
    ResetNamesCache

    Set sh = Wksh()

    'Admin tables: table N carries the levels 1 to N in its first N columns
    For counter = 1 To 4
        Set Lo = sh.ListObjects(TableNameAt(counter - 1))
        Set hRng = Lo.HeaderRowRange

        ReDim headerValues(1 To 1, 1 To counter)
        For counter2 = 1 To counter
            rawName = "adm" & counter2 & "_name"
            headerValues(1, counter2) = IIf(rawNames, rawName, GeoNames(rawName))
        Next

        hRng.Cells(1, 1).Resize(1, counter).Value = headerValues
    Next

    TranslateHFHeaders rawNames

    SetTranslatedFlag (Not rawNames)
    ResetNamesCache
End Sub

'@label:translate-hf-headers
'@sub-title Write the level headers of T_HF, each one resolved by name.
'@details
'The health facility table holds the level columns in an order that is not
'fixed and can carry a column this loop has nothing for. Each header is found
'by the label it holds now, so the two directions are symmetric and an extra
'column in the table shifts nothing.
'@param rawNames Boolean. True to write the raw column names, False to write the stored labels.
Private Sub TranslateHFHeaders(ByVal rawNames As Boolean)
    Dim hRng As Range
    Dim counter As Long
    Dim rawName As String
    Dim currentName As String
    Dim colIdx As Long

    Set hRng = Wksh().ListObjects("T_" & HF).HeaderRowRange

    For counter = 1 To 5
        rawName = LevelRawName(counter)
        currentName = IIf(rawNames, GeoNames(rawName), rawName)
        colIdx = ColumnIndex(currentName, hRng)

        If colIdx > 0 Then
            hRng.Cells(1, colIdx).Value = IIf(rawNames, rawName, GeoNames(rawName))
        End If
    Next
End Sub

'@section Clear
'===============================================================================

'@label:clear
'@sub-title Clear all geobase data.
'@details
'Puts the raw column names back on the table headers, empties the data body of
'every ListObject on the geobase worksheet, and sets the data flag to "empty".
'
'The header revert is what lets a caller clear and stop. EventsDesignerAdvanced
'calls this on its own, and a header left reading "Province" makes every column
'lookup in the class raise afterwards.
'
'Values go and formats stay: the population columns keep their number format
'and the validation of every column survives an import.
'@export
Public Sub Clear()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Translate rawNames:=True

    Set sh = Wksh()

    For Each Lo In sh.ListObjects
        If Lo.ListRows.Count > 0 Then Lo.DataBodyRange.ClearContents
    Next

    SetGeoUpdatedFlag FLAG_EMPTY
    ResetNamesCache
    ResetLevelMemo
End Sub

'@label:clear-historic
'@sub-title Clear historic entries.
'@details
'Empties the data body of the historic ListObjects named by the scope.
'@param scope Optional Byte. GeoScope value: GeoScopeAdmin (0) for T_HISTOGEO, GeoScopeHF (1) for T_HISTOHF, GeoScopeBoth (3) for both. Defaults to GeoScopeBoth.
'@export
Public Sub ClearHistoric(Optional ByVal scope As Byte = GeoScopeBoth)
    If (scope = GeoScopeAdmin) Or (scope = GeoScopeBoth) Then
        ClearHistoricTable "T_" & HISTOGEO
    End If

    If (scope = GeoScopeHF) Or (scope = GeoScopeBoth) Then
        ClearHistoricTable "T_" & HISTOHF
    End If
End Sub

'@label:clear-historic-table
'@sub-title Empty the data body of one historic table.
'@param loName String. The ListObject name.
Private Sub ClearHistoricTable(ByVal loName As String)
    Dim Lo As ListObject

    Set Lo = Wksh().ListObjects(loName)
    If Lo.ListRows.Count = 0 Then Exit Sub
    Lo.DataBodyRange.ClearContents
End Sub

'@section Import
'===============================================================================

'@label:import
'@sub-title Import geobase data from an external workbook.
'@details
'Imports geographic data from the supplied workbook into the geobase worksheet.
'Each worksheet whose name matches a known table name (ADM1 through ADM4, HF,
'NAMES, HISTOGEO, HISTOHF, METADATA) is imported column by column into the
'corresponding ListObject. Concat formulas are rewritten over the whole data
'column. When histoOnly is True only historic tables are refreshed and the
'main admin/HF data is left intact.
'
'A source workbook carrying a Metadata sheet also carries the hidden names
'ExportToFile wrote there, including the language code, and those are read
'back before the flags are set. The geobase is marked as carrying data only
'when the ADM1 table came across, so a workbook holding none of the expected
'sheets leaves the flag alone.
'@param fromWkb Workbook. Source workbook containing geo data sheets.
'@param histoOnly Optional Boolean. True to import only historic data. Defaults to False.
'@export
Public Sub Import(ByVal fromWkb As Workbook, _
                   Optional ByVal histoOnly As Boolean = False)

    Dim impsh As Worksheet
    Dim metash As Worksheet
    Dim admNames As BetterArray
    Dim importedAdmin1 As Boolean

    Set admNames = New BetterArray
    admNames.Push HISTOGEO, HISTOHF, METADATA

    If histoOnly Then
        ClearHistoric
    Else
        Clear
        admNames.Push ADMIN1, ADMIN2, ADMIN3, ADMIN4, HF, NAMES
    End If

    Translate rawNames:=True

    Set metash = Nothing

    For Each impsh In fromWkb.Worksheets
        If admNames.Includes(impsh.Name) Then
            ImportOneTable impsh
            If impsh.Name = ADMIN1 Then importedAdmin1 = True
            If impsh.Name = METADATA Then Set metash = impsh
        End If
    Next

    'The hidden names ExportToFile wrote travel on the Metadata sheet
    If Not metash Is Nothing Then
        GeoHiddenStore().ImportNames metash
        ResetFlagsCache
        'The headers are raw at this point, so the flag says so whatever the
        'source workbook carried.
        SetTranslatedFlag False
    End If

    GeoHiddenStore().SetValue HN_GEO_NAME, Replace(fromWkb.Name, ".xlsx", vbNullString)
    If importedAdmin1 Then SetGeoUpdatedFlag FLAG_UPDATED

    ResetNamesCache
    Translate rawNames:=False

    UpdateMetadata
    ResetNamesCache

    'Every level answered before this import belongs to the geobase that has
    'just been replaced. The historic-only path lands here too, and it costs
    'that path one rebuild of whatever the cascade asks for next.
    ResetLevelMemo
End Sub

'@label:import-one-table
'@sub-title Copy one source worksheet into its geobase table.
'@details
'The first row of the used range of the source is its header row, and the data
'runs from the row under it to the last row of that used range. A source sheet
'whose first row is blank therefore imports every record it holds.
'
'The destination row numbers come from the header row of the ListObject, so a
'table that does not start on row 1 lands in the right place.
'
'The concat column is written as a formula, calculated, and then replaced by the
'values it answered. The geo worksheet therefore carries no live formula once an
'import has run. A table filled by hand afterwards gets no concat for the rows
'added, because a column of constants is no longer a calculated column.
'@param impsh Worksheet. The source worksheet, named after the table it feeds.
Private Sub ImportOneTable(ByVal impsh As Worksheet)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim usedRng As Range
    Dim sourceHeaderRow As Long
    Dim sourceLastRow As Long
    Dim dataRows As Long
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim lastTableRow As Long
    Dim concatCol As Long
    Dim counter As Long
    Dim colName As String
    Dim colIdx As Long
    Dim concatFormula As String
    Dim concatRng As Range

    Set sh = Wksh()
    Set Lo = sh.ListObjects("T_" & impsh.Name)
    Set hRng = Lo.HeaderRowRange
    Set usedRng = impsh.UsedRange

    sourceHeaderRow = usedRng.Row
    sourceLastRow = usedRng.Row + usedRng.Rows.Count - 1
    dataRows = sourceLastRow - sourceHeaderRow
    If dataRows < 1 Then Exit Sub

    headerRow = hRng.Row
    firstDataRow = headerRow + 1
    lastDataRow = firstDataRow + dataRows - 1
    lastTableRow = Lo.Range.Row + Lo.Range.Rows.Count - 1
    concatCol = hRng.Column + hRng.Cells.Count - 1

    'Copy every source column whose header names a column of the table
    counter = 1
    Do While LenB(CellText(impsh.Cells(sourceHeaderRow, counter).Value)) > 0
        colName = CellText(impsh.Cells(sourceHeaderRow, counter).Value)
        colIdx = ColumnIndex(colName, hRng, inRange:=False)

        If colIdx > 0 Then
            sh.Range(sh.Cells(firstDataRow, colIdx), _
                     sh.Cells(lastDataRow, colIdx)).Value = _
                impsh.Range(impsh.Cells(sourceHeaderRow + 1, counter), _
                            impsh.Cells(sourceLastRow, counter)).Value
        End If

        counter = counter + 1
    Loop

    concatFormula = ConcatFormulaFor(Lo)

    'Clear the concat column over the table as it stands, then grow the table
    If LenB(concatFormula) > 0 And lastTableRow >= firstDataRow Then
        sh.Range(sh.Cells(firstDataRow, concatCol), _
                 sh.Cells(lastTableRow, concatCol)).ClearContents
    End If

    Lo.Resize sh.Range(sh.Cells(headerRow, hRng.Column), _
                       sh.Cells(lastDataRow, concatCol))


    '---------------------------------------------------------------------------
    'The formula goes over the whole data column, because a ListObject fills a
    'calculated column down on its own only while
    'Application.AutoCorrect.AutoFillFormulasInLists is on and a user can switch
    'that off. It is then calculated and replaced by what it answered, so the geo
    'worksheet ends with one value per row and no live formula.
    '
    'WHY THE EXPLICIT Calculate: every walk that imports a geobase applies a busy
    'state, and that state asks for manual calculation. Reading Value straight
    'after writing the formula would answer unevaluated cells and store zeros
    'over the whole geobase. Calculate on the range itself evaluates it whatever
    'mode the application is in, and the concat formula reads only constants
    'written into the same rows above, so the range is the whole dependency.
    If LenB(concatFormula) > 0 Then
        Set concatRng = sh.Range(sh.Cells(firstDataRow, concatCol), _
                                 sh.Cells(lastDataRow, concatCol))
        concatRng.Formula = concatFormula
        concatRng.Calculate
        concatRng.Value = concatRng.Value
    End If
End Sub

'@label:concat-formula-for
'@sub-title Build the concat formula of one geobase table.
'@details
'The concat column of a table joins the level columns above it with a pipe, so
'a filtered lookup can match one full path. A table with no concat column
'answers an empty string.
'@param Lo ListObject. The geobase table.
'@return String. The formula for the first data row, or an empty string.
Private Function ConcatFormulaFor(ByVal Lo As ListObject) As String
    Dim rowRng As Range
    Dim sep As String
    Dim parts As Long
    Dim counter As Long
    Dim formulaText As String

    Select Case Lo.Name
    Case "T_ADM1"
        parts = 1
    Case "T_ADM2"
        parts = 2
    Case "T_ADM3"
        parts = 3
    Case "T_ADM4", "T_HF"
        parts = 4
    Case Else
        Exit Function
    End Select

    sep = " & " & Chr(34) & " | " & Chr(34) & " & "
    Set rowRng = Lo.HeaderRowRange.Offset(1)

    formulaText = "=" & rowRng.Cells(1, 1).Address(RowAbsolute:=False)
    For counter = 2 To parts
        formulaText = formulaText & sep & _
                      rowRng.Cells(1, counter).Address(RowAbsolute:=False)
    Next

    ConcatFormulaFor = formulaText
End Function

'@section Export
'===============================================================================

'@label:export-to-file
'@sub-title Export geobase tables to a workbook as flat data.
'@details
'Exports geobase tables to the target workbook with one sheet per table and a
'Metadata sheet. Headers go back to raw names before the export and take their
'labels again afterwards, and they do so whichever way the export ends.
'HiddenNames entries of the geo worksheet are written onto the Metadata sheet,
'and Import reads them back from there.
'When onlyHistory is True only the HISTOGEO and HISTOHF tables are exported.
'@param outputWkb Workbook. Target workbook to receive exported sheets.
'@param onlyHistory Optional Boolean. True to export only historic tables. Defaults to False.
'@export
Public Sub ExportToFile(ByVal outputWkb As Workbook, _
                         Optional ByVal onlyHistory As Boolean = False)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim sheetName As String
    Dim loName As String
    Dim counter As Long
    Dim wasTranslated As Boolean
    Dim atStep As String
    Dim priorCalc As Long
    Dim calcHeld As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    this.lastFailure = vbNullString

    'CALCULATION IS HELD HERE, BY THIS ROUTINE
    '---------------------------------------------------------------------------
    'Adding a worksheet recalculates. ImportOneTable now leaves the concat column
    'as values, so a geobase imported by the current code brings no formulas with
    'it -- but a linelist filled before that change carries one concat formula per
    'row of every table, hundreds of thousands of them on a real geobase, and this
    'routine has to be safe over those files too.
    'ApplicationState.ApplyBusyState already asks for manual, but it asks inside
    'an On Error Resume Next and reports success either way, so a caller cannot
    'know the mode it got. This holds the mode itself and does not depend on that.
    priorCalc = Application.Calculation
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
    calcHeld = (Application.Calculation = xlCalculationManual)

    'The handler covers the whole routine. A failure before the headers are
    'reverted leaves wasTranslated False, so the restore below does nothing,
    'which is what an untouched header row wants.
    On Error GoTo RestoreHeaders

    atStep = "reading the geo worksheet"
    Set sh = Wksh()

    atStep = "listing the tables to export"
    If onlyHistory Then
        Set tabl = New BetterArray
        tabl.Push "T_" & HISTOGEO, "T_" & HISTOHF
    Else
        Set tabl = TableNames()
    End If

    'BEST EFFORT, BECAUSE THE HEADERS ARE NOT THE EXPORT
    '---------------------------------------------------------------------------
    'What this step changes is which words sit on the header row of the admin
    'tables -- adm1_name, or the label the geobase was translated into. The
    'export is the DATA underneath them, and a file whose headers could not be
    'put back to their raw names is a file worth writing.
    'It refuses on a linelist older than the hidden name store: the store tracks
    'a name only when it is hidden, that file keeps its geo names visible, and
    'SetValue therefore finds nothing to write. Failing the whole geobase export
    'over the wording of a header row is the wrong trade.
    'The failure is not swallowed. It is written into the linelist''s own log
    'and kept in lastFailure, so a caller and a user both see that the headers
    'were left as they were.
    atStep = "putting the raw names back on the headers"
    TryTranslate rawNames:=True, wasTranslated:=wasTranslated

    With outputWkb
        atStep = "naming the Metadata sheet"
        .Worksheets(1).Name = METADATASHEET

        'EVERY SHEET IS ADDED BEFORE ANY DATA IS WRITTEN
        '-----------------------------------------------------------------------
        'Adding a worksheet recalculates. Adding one between block writes pays
        'that over everything written so far, and on a geobase of a few hundred
        'thousand rows the add is where the export used to stop.
        sheetName = METADATASHEET
        For counter = tabl.LowerBound To tabl.UpperBound
            loName = tabl.Item(counter)

            atStep = "adding the sheet " & Replace(loName, "T_", vbNullString) & _
                     " after " & sheetName & " (the workbook holds " & _
                     SheetNamesOf(outputWkb) & ")"
            .Worksheets.Add(After:=.Worksheets(sheetName)).Name = _
                Replace(loName, "T_", vbNullString)
            sheetName = Replace(loName, "T_", vbNullString)
        Next

        'THE DATA GOES OUT THROUGH BetterArray, ONE COLUMN AT A TIME
        '-----------------------------------------------------------------------
        'BetterArray reads a range into memory and writes it back, and it sizes
        'the destination from what it read, so nothing here counts rows.
        'ExportOneTable does one column per pass. A table read whole holds the
        'same cells three times at once, and on T_ADM4 that is a few million
        'cells of each. A column is a seventh of that.
        ExportOneTable sh.ListObjects("T_" & METADATA), _
                       .Worksheets(METADATASHEET), atStep

        For counter = tabl.LowerBound To tabl.UpperBound
            loName = tabl.Item(counter)
            sheetName = Replace(loName, "T_", vbNullString)
            ExportOneTable sh.ListObjects(loName), .Worksheets(sheetName), atStep
        Next

        'The hidden names go onto the Metadata sheet, and Import reads them
        atStep = "writing the hidden names onto the Metadata sheet"
        GeoHiddenStore().ExportNames .Worksheets(METADATASHEET)
    End With

    atStep = "putting the labels back on the headers"
    If wasTranslated Then Translate rawNames:=False

    On Error Resume Next
    Application.Calculation = priorCalc
    On Error GoTo 0
    Exit Sub

RestoreHeaders:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    'Kept in words, because the raise below loses them on the way out of this
    'class: the caller reads the name of the method and nothing of what refused.
    'The calculation mode rides along, because a mode that would not go manual
    'is the difference between a geobase that exports and one that does not.
    this.lastFailure = "geobase export failed " & atStep & " - error " & _
                       errNumber & " from " & errSource & ": " & errMessage & _
                       " (calculation " & IIf(calcHeld, "held manual", _
                       "NOT manual, mode " & Application.Calculation) & ")"
    Resume RaiseAfterRestore

RaiseAfterRestore:
    On Error Resume Next
    If wasTranslated Then Translate rawNames:=False
    Application.Calculation = priorCalc
    On Error GoTo 0
    Err.Raise errNumber, errSource, errMessage
End Sub

'@label:export-one-table
'@sub-title Write one geobase table onto a sheet, a column at a time.
'@details
'The header row goes out with the data, because a ListColumn range carries its
'header, and Import reads the first used row of the sheet as the header row.
'The destination starts at A1 and each column lands under its own position in
'the table, so the sheet comes out in the order of the table.
'BetterArray sizes the destination from what it read, so no row count is worked
'out here, and an empty table writes its header row and nothing else.
'The step is written back to the caller on every column, so a failure names the
'column it stopped on rather than the table.
'@param Lo ListObject. The geobase table to write out.
'@param outsh Worksheet. The sheet to write it onto.
'@param atStep String. Updated in place with the column being written.
Private Sub ExportOneTable(ByVal Lo As ListObject, _
                            ByVal outsh As Worksheet, _
                            ByRef atStep As String)

    Dim arr As BetterArray
    Dim colIndex As Long
    Dim colCount As Long
    Dim rowCount As Long

    colCount = Lo.ListColumns.Count
    rowCount = Lo.Range.Rows.Count

    For colIndex = 1 To colCount
        atStep = "writing column " & colIndex & " of " & colCount & ", " & _
                 Lo.ListColumns(colIndex).Name & ", of " & Lo.Name & _
                 " onto " & outsh.Name & " (" & rowCount & " rows)"

        'A new BetterArray per column, so no bound or type of the column before
        'it is still in place when this one is read.
        Set arr = New BetterArray
        arr.FromExcelRange Lo.ListColumns(colIndex).Range
        arr.ToExcelRange outsh.Cells(1, colIndex)
    Next

End Sub

'@label:sheet-names-of
'@sub-title The worksheet names of a workbook with their used rows, on one line.
'@details
'Written for the failure account of ExportToFile. Two questions a reader has
'when a sheet will not be added: whether the name was already taken, and how
'much had already been written into the workbook. Each sheet is reported as
'name(used rows), so both are answered by one line.
'@param wkb Workbook. The workbook to list.
'@return String. The sheets separated by commas, or a note when they cannot be read.
Private Function SheetNamesOf(ByVal wkb As Workbook) As String
    Dim sh As Worksheet
    Dim names As String
    Dim usedRows As String

    On Error Resume Next
    For Each sh In wkb.Worksheets
        usedRows = "?"
        usedRows = CStr(sh.UsedRange.Rows.Count)
        If LenB(names) > 0 Then names = names & ", "
        names = names & sh.Name & "(" & usedRows & ")"
    Next
    On Error GoTo 0

    If LenB(names) = 0 Then names = "no readable sheet names"
    SheetNamesOf = names
End Function

'@label:last-failure
'@prop-title What the last export of this geobase failed on
'@details
'ExportToFile raises, and the caller reads a description that names the method
'and nothing else, because the error crosses out of this class. The step, the
'number, the source and the text are kept here before the raise, so the caller
'can log what actually refused. ExportToFile empties it on entry.
'@return String. The step, error number, source and text, or vbNullString.
'@export
Public Property Get LastFailure() As String
    LastFailure = this.lastFailure
End Property

'@label:export-to-wkb
'@sub-title Export the geobase worksheet structure to a workbook.
'@details
'Copies the entire geobase worksheet to the target workbook, preserving
'ListObjects, named ranges, formulas, pcode references, concat references,
'and metadata. Creates a hidden Metadata sheet with creation metadata: the
'linelist name, the setup name, the language, the timestamp, the platform the
'build ran on and the version of the designer that ran it. RNG_PastingGeoCol is
'exported as a cell-based named range. The hidden names of the geo worksheet
'go with the sheet, and the five workbook-scoped level labels are written into
'the target workbook so CrossTable's spatial dropdown resolves them there.
'Raises when sheets with conflicting names already exist in the target
'workbook.
'@param outputWkb Workbook. Target workbook.
'@param llName Optional String. Linelist name for metadata. Defaults to vbNullString.
'@param setupPath Optional String. Setup file path for metadata. Defaults to vbNullString.
'@throws ProjectError.ElementShouldNotExists When target sheets already exist.
'@export
Public Sub ExportToWkb(ByVal outputWkb As Workbook, _
                        Optional ByVal llName As String = vbNullString, _
                        Optional ByVal setupPath As String = vbNullString)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As ListObject
    Dim expAddress As String
    Dim expRange As Range
    Dim rng As Range
    Dim metaKeys As BetterArray
    Dim metaValues As BetterArray
    Dim setupName As String
    Dim wasTranslated As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    Set sh = Wksh()

    'Check that target sheets do not already exist
    Set outsh = Nothing
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(sh.Name)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & sh.Name & " in the workbook"
    End If

    Set outsh = Nothing
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & METADATASHEET & " in the workbook"
    End If

    wasTranslated = IsTranslated()
    Translate rawNames:=False

    On Error GoTo RestoreHeaders

    'Create the geo sheet and copy ListObjects. The placement is explicit: a
    'bare Add puts the new sheet before whatever sheet is active, which used to
    'land Geo and Metadata at or near position 1. Every other exporter appends,
    'and the final sheet order stopped depending on the active sheet.
    outputWkb.Worksheets.Add( _
        After:=outputWkb.Worksheets(outputWkb.Worksheets.Count) _
    ).Name = sh.Name
    Set outsh = outputWkb.Worksheets(sh.Name)

    For Each Lo In sh.ListObjects
        expAddress = Lo.Range.Address
        Set expRange = outsh.Range(expAddress)
        expRange.Value = Lo.Range.Value

        With outsh
            .ListObjects.Add( _
                SourceType:=xlSrcRange, _
                Source:=expRange, _
                XlListObjectHasHeaders:=xlYes _
            ).Name = Lo.Name
            .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
        End With
    Next

    outsh.UsedRange.EntireColumn.Locked = False
    outsh.UsedRange.EntireColumn.AutoFit

    'The paste anchor stays a real cell
    expAddress = sh.Range(RNG_PASTING_GEO_COL).Address
    Set expRange = outsh.Range(expAddress)
    expRange.Name = RNG_PASTING_GEO_COL
    expRange.Value = sh.Range(RNG_PASTING_GEO_COL).Value
    expRange.Locked = True

    'The hidden names of the geo worksheet travel with the sheet
    GeoHiddenStore().ExportNames outsh

    'The five level labels are workbook-scoped, so they are written into the
    'target workbook. They are written after the translate, so they carry the
    'labels the exported headers carry.
    ExportLevelNames outputWkb

    'Create metadata sheet, appended for the same reason as the geo sheet above
    outputWkb.Worksheets.Add( _
        After:=outputWkb.Worksheets(outputWkb.Worksheets.Count) _
    ).Name = METADATASHEET
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)

    On Error Resume Next
    setupName = Split(setupPath, Application.PathSeparator)( _
        UBound(Split(setupPath, Application.PathSeparator)))
    On Error GoTo RestoreHeaders

    Set metaKeys = New BetterArray
    Set metaValues = New BetterArray

    'The platform and the designer version sit beside the four rows that were
    'already here. A linelist is read on a machine that is not the one that
    'built it, and the two platforms do not behave the same, so the file has to
    'say which one made it and which designer did it.
    'The anchor below is placed off metaKeys.Length, so both blocks move down
    'on their own and nothing here counts rows.
    metaKeys.Push "variable", "language", "linelist_creation_date", _
                  "linelist_creation_name", "linelist_creation_os", _
                  "used_setup_name", "used_designer_name", _
                  "used_designer_version"

    metaValues.Push "value", GeoHiddenStore().ValueAsString(HN_META_LANG), _
                    Format(Now(), "yyyy-mm-dd Hh:Nn"), _
                    llName, PlatformTag(), setupName, ThisWorkbook.Name, _
                    DesignerVersion()

    With outsh
        metaKeys.ToExcelRange .Cells(1, 1)
        metaValues.ToExcelRange .Cells(1, 2)

        Set expRange = .Cells(metaKeys.Length + 1, 1)
        expRange.Name = "RNG_Meta_Anchor"
        Set expRange = .Range(expRange, expRange.Offset(rng.Rows.Count - 1, 1))
        expRange.Value = rng.Value
        .Visible = xlSheetVeryHidden
    End With

    'Add workbook-level named references for pcodes, concat, population, and historic
    With outputWkb.Names
        .Add Name:="adm1_pcode", RefersToR1C1:="=T_ADM1[adm1_pcode]"
        .Add Name:="adm2_pcode", RefersToR1C1:="=T_ADM2[adm2_pcode]"
        .Add Name:="adm3_pcode", RefersToR1C1:="=T_ADM3[adm3_pcode]"
        .Add Name:="adm4_pcode", RefersToR1C1:="=T_ADM4[adm4_pcode]"
        .Add Name:="hf_pcode", RefersToR1C1:="=T_HF[hf_pcode]"

        .Add Name:="adm1_concat", RefersToR1C1:="=T_ADM1[adm1_concat]"
        .Add Name:="adm2_concat", RefersToR1C1:="=T_ADM2[adm2_concat]"
        .Add Name:="adm3_concat", RefersToR1C1:="=T_ADM3[adm3_concat]"
        .Add Name:="adm4_concat", RefersToR1C1:="=T_ADM4[adm4_concat]"
        .Add Name:="hf_concat", RefersToR1C1:="=T_HF[hf_concat]"

        .Add Name:="adm1_pop", RefersToR1C1:="=T_ADM1[adm1_pop]"
        .Add Name:="adm2_pop", RefersToR1C1:="=T_ADM2[adm2_pop]"
        .Add Name:="adm3_pop", RefersToR1C1:="=T_ADM3[adm3_pop]"
        .Add Name:="adm4_pop", RefersToR1C1:="=T_ADM4[adm4_pop]"

        .Add Name:="histo_geo", RefersToR1C1:="=T_HISTOGEO[HistoGeo]"
        .Add Name:="histo_hf", RefersToR1C1:="=T_HISTOHF[HistoFacility]"
    End With

    Exit Sub

RestoreHeaders:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    Resume RaiseAfterRestore

RaiseAfterRestore:
    On Error Resume Next
    If Not wasTranslated Then Translate rawNames:=True
    On Error GoTo 0
    Err.Raise errNumber, errSource, errMessage
End Sub

'@label:platform-tag
'@sub-title The platform the linelist was built on, on one tag
'@details
'A linelist is opened on machines that are not the one that built it, and the
'two platforms do not behave the same: a geobase export Windows accepts has
'refused on a Mac. The metadata sheet therefore names the platform of the
'build. The name and the bitness come from the compile constants, so they are
'the build that is running and cannot be misread. The Excel version comes from
'the application.
'@return String. Something of the shape "mac-64 excel-16.90".
Private Function PlatformTag() As String
    Dim osName As String
    Dim bits As String

    #If Mac Then
        osName = "mac"
    #Else
        osName = "win"
    #End If

    #If Win64 Then
        bits = "64"
    #Else
        bits = "32"
    #End If

    PlatformTag = osName & "-" & bits & " excel-" & Application.Version
End Function

'@label:designer-version
'@sub-title The version of the designer that built the linelist
'@details
'The version is a workbook-scoped hidden name on the designer.
'DesignerPreparation.Prepare writes it from the RNG_Version value name the
'owner keeps on the Dev worksheet, so a designer that was never initialised
'with a version there does not carry it. That is the ordinary case rather than
'a fault, and it answers "(not found)". A name that is present but empty
'answers the same, because an empty cell on the metadata sheet says nothing to
'the person reading it.
'@return String. The version, or "(not found)".
Private Function DesignerVersion() As String
    Dim store As HiddenNames
    Dim versionText As String

    DesignerVersion = VERSION_NOT_FOUND

    'The designer is the workbook this code runs from, the same one
    'used_designer_name is taken from. A workbook that refuses a name store
    'leaves the answer at "(not found)" rather than taking down the build.
    On Error Resume Next
    Set store = HiddenNames.Create(ThisWorkbook)
    On Error GoTo 0
    If store Is Nothing Then Exit Function

    If Not store.HasName(HN_DESIGNER_VERSION) Then Exit Function

    versionText = Trim$(store.ValueAsString(HN_DESIGNER_VERSION))
    If LenB(versionText) = 0 Then Exit Function

    DesignerVersion = versionText
End Function

'@label:export-level-names
'@sub-title Write the five level labels into another workbook.
'@details
'The five labels are workbook-scoped hidden names, so a sheet copy leaves them
'behind. Each one is written by name into the target workbook.
'@param outputWkb Workbook. The workbook to write them into.
Private Sub ExportLevelNames(ByVal outputWkb As Workbook)
    Dim destination As HiddenNames
    Dim counter As Long
    Dim nameId As String
    Dim label As String

    Set destination = HiddenNames.Create(outputWkb)

    For counter = 1 To 5
        nameId = LevelNameId(LevelRawName(counter))
        label = GeoNames(LevelRawName(counter))
        destination.EnsureName nameId, label, HiddenNameTypeString
        destination.SetValue nameId, label
    Next
End Sub

'@section Update
'===============================================================================

'@label:update
'@sub-title Update linelist geo column headers and dropdowns.
'@details
'Walks the HList worksheets of the parent workbook, finds the columns
'controlled by geo1 through geo4 and hf, and writes the level labels into
'their header cells and into the matching cells of the print sheet. Then
'refreshes the admin1 dropdown list on the __dropdown_lists sheet.
'
'A sheet the walk cannot finish leaves its name in SkippedSheets and the walk
'carries on with the next one. Every sheet that is unprotected is protected
'again, whichever way its own pass ends.
'@param pass Passwords. Password manager for sheet protection and unprotection.
'@export
Public Sub Update(ByVal pass As Passwords)

    Const DROPDOWNSHEET As String = "__dropdown_lists"

    Dim sh As Worksheet
    Dim dropSh As Worksheet
    Dim wb As Workbook
    Dim geoData As BetterArray
    Dim drop As DropdownLists
    Dim shHn As HiddenNames

    Set wb = Wksh.Parent
    Set this.skipped = New BetterArray

    For Each sh In wb.Worksheets
        Set shHn = HiddenNames.Create(sh)

        If shHn.ValueAsString(SHEET_TAG) = "HList" Then
            UpdateOneHList sh, wb, shHn, pass
        End If
    Next

    'The dropdown sheet gets its own variable, so the guard below reads the
    'answer to its own question.
    Set dropSh = Nothing
    On Error Resume Next
    Set dropSh = wb.Worksheets(DROPDOWNSHEET)
    On Error GoTo 0

    If dropSh Is Nothing Then Exit Sub

    Set drop = DropdownLists.Create(dropSh)
    Set geoData = GeoLevel(LevelAdmin1)
    drop.Update geoData, "admin1"
End Sub

'@label:update-one-hlist
'@sub-title Write the level labels into the headers of one HList sheet.
'@details
'The variable a column belongs to is the value of the _START cell of its
'column, and its control value is a hidden name built from that variable name.
'VarWriter writes that name as "<variable> -- control".
'
'The label goes into the row above the data, on the sheet and on its print
'partner. A sheet with no print partner keeps its own headers updated and is
'reported through SkippedSheets.
'@param sh Worksheet. The HList worksheet.
'@param wb Workbook. The workbook holding the sheet.
'@param shHn HiddenNames. The hidden name store of that sheet.
'@param pass Passwords. Password manager for sheet protection.
Private Sub UpdateOneHList(ByVal sh As Worksheet, ByVal wb As Workbook, _
                            ByVal shHn As HiddenNames, ByVal pass As Passwords)
    Dim cellRng As Range
    Dim cellPrintRng As Range
    Dim printsh As Worksheet
    Dim tableName As String
    Dim controlValue As String
    Dim levelLabel As String
    Dim recorded As Boolean

    tableName = shHn.ValueAsString("table_name")
    If LenB(tableName) = 0 Then
        this.skipped.Push sh.Name
        Exit Sub
    End If

    pass.UnProtect sh.Name

    On Error GoTo SkipSheet

    Set cellRng = sh.Range(tableName & "_START")

    Set printsh = Nothing
    On Error Resume Next
    Set printsh = wb.Worksheets("print_" & sh.Name)
    On Error GoTo SkipSheet

    If printsh Is Nothing Then
        this.skipped.Push sh.Name
        recorded = True
    Else
        Set cellPrintRng = printsh.Range(cellRng.Address)
    End If

    Do While Not IsEmpty(cellRng)
        controlValue = shHn.ValueAsString(cellRng.Value & " -- control")
        levelLabel = LevelLabelForControl(controlValue)

        If LenB(levelLabel) > 0 Then
            cellRng.Offset(-1).Value = levelLabel
            If Not cellPrintRng Is Nothing Then cellPrintRng.Offset(-1).Value = levelLabel
        End If

        Set cellRng = cellRng.Offset(, 1)
        If Not cellPrintRng Is Nothing Then Set cellPrintRng = cellPrintRng.Offset(, 1)
    Loop

    pass.Protect sh.Name
    Exit Sub

SkipSheet:
    Resume ProtectAfterFail

ProtectAfterFail:
    On Error Resume Next
    If Not recorded Then this.skipped.Push sh.Name
    pass.Protect sh.Name
    On Error GoTo 0
End Sub

'@label:level-label-for-control
'@sub-title The level label a control value asks for.
'@param controlValue String. The control value of a dictionary variable.
'@return String. The label of the level, or an empty string for a control that is not geographic.
Private Function LevelLabelForControl(ByVal controlValue As String) As String
    Select Case controlValue
    Case "geo1"
        LevelLabelForControl = GeoNames("adm1_name")
    Case "geo2"
        LevelLabelForControl = GeoNames("adm2_name")
    Case "geo3"
        LevelLabelForControl = GeoNames("adm3_name")
    Case "geo4"
        LevelLabelForControl = GeoNames("adm4_name")
    Case "hf"
        LevelLabelForControl = GeoNames("hf_name")
    End Select
End Function

'@label:skipped-sheets
'@prop-title The sheets the last Update could not finish.
'@details
'Holds the name of every HList worksheet Update left part way through, which
'today means a sheet with no table_name hidden name, a sheet whose _START cell
'cannot be resolved, and a sheet with no print_ partner. An empty list means
'every HList sheet was updated in full.
'@return BetterArray. Worksheet names, empty when Update has not run.
'@export
Public Property Get SkippedSheets() As BetterArray
    If this.skipped Is Nothing Then Set this.skipped = New BetterArray
    Set SkippedSheets = this.skipped
End Property

'@label:update-dict
'@sub-title Update dictionary labels for geo-related variables.
'@details
'Scans the dictionary control column for geo1 through geo4 and hf entries and
'overwrites the corresponding main label cells with the level labels of the
'geobase. This keeps the dictionary in step after a geobase import or a
'translation change.
'@param dict LLdictionary. The linelist dictionary whose labels should be updated.
'@export
Public Sub UpdateDict(ByVal dict As LLdictionary)
    Dim controlRng As Range
    Dim mainlabRng As Range
    Dim counter As Long
    Dim levelLabel As String

    Set controlRng = dict.DataRange("control")
    Set mainlabRng = dict.DataRange("main label")

    For counter = 1 To controlRng.Rows.Count
        levelLabel = LevelLabelForControl(CellText(controlRng.Cells(counter, 1).Value))
        If LenB(levelLabel) > 0 Then mainlabRng.Cells(counter, 1).Value = levelLabel
    Next
End Sub

'@section Historic
'===============================================================================

'@label:update-historic
'@sub-title Add a new entry to the historic data.
'@details
'Appends a value to the appropriate historic ListObject (T_HISTOGEO or
'T_HISTOHF), grows the table to include the new row, then removes duplicates
'and sorts the table as it now stands. The range is re-read after the resize,
'so the row just written is de-duplicated and placed with the rest.
'@param newValue String. The value to add to the historic list.
'@param scope Optional Byte. GeoScope value: GeoScopeAdmin (0) for T_HISTOGEO, GeoScopeHF (1) for T_HISTOHF. Defaults to GeoScopeAdmin.
'@export
Public Sub UpdateHistoric(ByVal newValue As String, _
                           Optional ByVal scope As Byte = GeoScopeAdmin)
    Dim sh As Worksheet
    Dim loName As String
    Dim LoRng As Range
    Dim cellRng As Range
    Dim Lo As ListObject

    Set sh = Wksh()
    loName = IIf(scope = GeoScopeHF, HISTOHF, HISTOGEO)
    loName = "T_" & loName

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range
    Set cellRng = LoRng.Cells(LoRng.Rows.Count + 1, 1)

    If Lo.ListRows.Count = 0 Then
        Set cellRng = LoRng.Cells(2, 1)
    ElseIf Lo.ListRows.Count = 1 Then
        If LenB(CellText(Lo.DataBodyRange.Cells(1, 1).Value)) = 0 Then
            Set cellRng = LoRng.Cells(2, 1)
        End If
    End If

    cellRng.Value = newValue
    Lo.Resize sh.Range(LoRng.Cells(1, 1), cellRng)

    'Re-read the table: the two calls below work on the row just written
    Set LoRng = Lo.Range
    LoRng.RemoveDuplicates Columns:=1, Header:=xlYes

    If Lo.ListRows.Count > 1 Then
        Lo.Range.Sort Key1:=Lo.ListColumns(1).Range, Header:=xlYes
    End If
End Sub

'@section AdminCode
'===============================================================================

'@label:admin-code
'@sub-title Resolve the admin code from a level label.
'@details
'Compares the supplied label against the label of each admin level and returns
'the matching raw code ("adm1" through "adm4").
'
'A label that matches no level answers an empty string. It used to answer
'"adm1", which reads on the sheet as a deliberate choice of the first level.
'Three states reach this arm and none of them is that choice: a freshly
'generated table still carrying its "select a level" prompt in the dropdown
'cell, a dropdown a user cleared with Delete, and a geobase re-imported under
'labels the cell never saw. The two call sites in EventLinelist that rewrite
'formulas refuse an empty answer outright, and the two that read fall back to
'PREVIOUS_ADM_<tabId>, the cell recording the level those formulas stand on.
'@param selectedAdmin String. The level label to resolve.
'@return String. Admin code ("adm1", "adm2", "adm3", or "adm4"), or an empty
'string when the label names no level.
'@export
Public Function AdminCode(ByVal selectedAdmin As String) As String
    Dim adm1Name As String
    Dim adm2Name As String
    Dim adm3Name As String
    Dim adm4Name As String

    adm1Name = GeoNames("adm1_name")
    adm2Name = GeoNames("adm2_name")
    adm3Name = GeoNames("adm3_name")
    adm4Name = GeoNames("adm4_name")

    Select Case selectedAdmin
    Case adm1Name
        AdminCode = "adm1"
    Case adm2Name
        AdminCode = "adm2"
    Case adm3Name
        AdminCode = "adm3"
    Case adm4Name
        AdminCode = "adm4"
    Case Else
        AdminCode = vbNullString
    End Select
End Function

'@label:has-no-data
'@prop-title Whether the geobase has no imported data.
'@details
'Answers True when the data flag reads "empty", which is the state a geobase
'is in before its first import and after Clear. The flag says nothing about
'the headers, so an export that reverts and re-applies the headers leaves this
'answer alone.
'@return Boolean. True when the geobase is empty.
'@export
Public Property Get HasNoData() As Boolean
    HasNoData = (GeoUpdatedFlag() = FLAG_EMPTY)
End Property

'@section Internal Helpers
'===============================================================================

'@label:update-metadata
'@sub-title Update the metadata sheet from the geo metadata table.
'@details
'Copies the current content of T_METADATA into the hidden Metadata worksheet
'at the RNG_Meta_Anchor position, after clearing whatever was written there
'before. A workbook with no Metadata sheet and a sheet with no anchor are both
'ordinary states and answer with no work. A failure of the write itself is
'raised.
Private Sub UpdateMetadata()
    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim metash As Worksheet
    Dim anchorRng As Range
    Dim rng As Range
    Dim expRng As Range
    Dim lastRow As Long

    Set sh = Wksh()

    Set metash = Nothing
    On Error Resume Next
    Set metash = sh.Parent.Worksheets(METADATASHEET)
    On Error GoTo 0
    If metash Is Nothing Then Exit Sub

    Set anchorRng = Nothing
    On Error Resume Next
    Set anchorRng = metash.Range("RNG_Meta_Anchor")
    On Error GoTo 0
    If anchorRng Is Nothing Then Exit Sub

    With metash
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        If lastRow >= anchorRng.Row Then
            Set rng = .Range(anchorRng, .Cells(lastRow, 2))
            rng.ClearContents
        End If

        Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)
        Set expRng = .Range(anchorRng, anchorRng.Offset(rng.Rows.Count - 1, 1))
        expRng.Value = rng.Value
    End With
End Sub

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate that the worksheet carries what the class reads.
'@details
'Verifies the presence of all nine required ListObjects (T_ADM1 through T_ADM4,
'T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF, T_METADATA), the four hidden names of
'the geo worksheet (RNG_GeoUpdated, RNG_GeoName, RNG_GeoLangCode,
'RNG_MetaLang), and the RNG_PastingGeoCol cell.
'
'The five level labels are ensured. A Geo worksheet built before they moved
'into HiddenNames carries them as cells: the cell is read once, the hidden
'name is written with that value, and no reader touches the cell again. A
'worksheet with no such cell gets five empty names, which UpdateLevelNames
'fills on the next Translate.
'@param sh Worksheet. The worksheet to validate.
'@param geoStore HiddenNames. The store over that worksheet.
'@param levelStore HiddenNames. The store over the workbook that owns it.
'@throws ProjectError.ElementNotFound When required tables, named ranges, or hidden names are missing.
Private Sub CheckRequirements(ByVal sh As Worksheet, ByVal geoStore As HiddenNames, _
                               ByVal levelStore As HiddenNames)
    Dim admTableNames As BetterArray

    Set admTableNames = New BetterArray
    admTableNames.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
                       "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF, _
                       "T_" & METADATA
    LoExists admTableNames, sh

    HiddenNameRequired geoStore, HN_GEO_NAME
    HiddenNameRequired geoStore, HN_GEO_UPDATED
    HiddenNameRequired geoStore, HN_GEO_LANG_CODE
    HiddenNameRequired geoStore, HN_META_LANG

    RangeExists sh, RNG_PASTING_GEO_COL

    EnsureLevelNames sh, levelStore
End Sub

'@label:ensure-level-names
'@sub-title Ensure the five level labels exist as workbook hidden names.
'@details
'A level whose hidden name is missing takes the value of the cell-based named
'range of the same identifier when the worksheet still carries one, and an
'empty value otherwise. That read is the whole migration off the cells.
'@param sh Worksheet. The geobase worksheet.
'@param levelStore HiddenNames. The store over the workbook that owns it.
Private Sub EnsureLevelNames(ByVal sh As Worksheet, ByVal levelStore As HiddenNames)
    Dim counter As Long
    Dim nameId As String
    Dim seedValue As String

    For counter = 1 To 5
        nameId = LevelNameId(LevelRawName(counter))

        If Not levelStore.HasName(nameId) Then
            seedValue = CellSeedValue(sh, nameId)
            levelStore.EnsureName nameId, seedValue, HiddenNameTypeString
        End If
    Next
End Sub

'@label:cell-seed-value
'@sub-title Read the value of a cell-based named range, when there is one.
'@param sh Worksheet. The worksheet to read from.
'@param rngName String. The named range to look for.
'@return String. The cell value, or an empty string when the range is absent or holds an error.
Private Function CellSeedValue(ByVal sh As Worksheet, ByVal rngName As String) As String
    Dim rng As Range

    Set rng = Nothing
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Function

    CellSeedValue = CellText(rng.Cells(1, 1).Value)
End Function

'@label:hidden-name-required
'@sub-title Verify that a hidden name exists in the store.
'@details
'Checks whether the specified name exists in the given HiddenNames store.
'If not found, raises an ElementNotFound error with a descriptive message.
'@param store HiddenNames. The hidden names manager to check.
'@param nameId String. The hidden name identifier to look for.
'@throws ProjectError.ElementNotFound When the hidden name does not exist.
Private Sub HiddenNameRequired(ByVal store As HiddenNames, ByVal nameId As String)
    If Not store.HasName(nameId) Then
        ThrowError ElementNotFound, _
            "Hidden name " & nameId & " does not exist on the geobase worksheet"
    End If
End Sub

'@label:range-exists
'@sub-title Verify that a named range exists on a worksheet.
'@details
'Attempts to resolve a named range on the given worksheet. If the range cannot
'be found, raises an ElementNotFound error with a descriptive message.
'@param sh Worksheet. The worksheet to check.
'@param rngName String. The named range to look for.
'@throws ProjectError.ElementNotFound When the named range does not exist on the worksheet.
Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then
        ThrowError ElementNotFound, _
            "Range " & rngName & " does not exist in worksheet " & sh.Name
    End If
End Sub

'@label:lo-exists
'@sub-title Verify that all listed ListObjects exist on a worksheet.
'@details
'Iterates over the supplied BetterArray of table names and checks that each
'ListObject can be found on the given worksheet. Raises an ElementNotFound
'error for the first missing table encountered.
'@param loNames BetterArray. List of ListObject names to verify.
'@param sh Worksheet. The worksheet to check.
'@throws ProjectError.ElementNotFound When a required ListObject does not exist on the worksheet.
Private Sub LoExists(ByVal loNames As BetterArray, ByVal sh As Worksheet)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String

    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        Set Lo = Nothing

        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0

        If Lo Is Nothing Then
            ThrowError ElementNotFound, _
                "Table " & lName & " does not exist in worksheet " & sh.Name
        End If
    Next
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception.
'@details
'Wraps Err.Raise to throw a structured error using a ProjectError error code
'and a descriptive message. The error number is passed as Long because
'ProjectError values exceed the Byte range (they start at 1001+).
'@param errNumber Long. The ProjectError error code to raise.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
