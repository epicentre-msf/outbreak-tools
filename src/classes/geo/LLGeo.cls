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
Implements ILLGeo

'@class LLGeo
'@description
'Manages the geobase worksheet that stores the four-level administrative
'hierarchy (admin1 through admin4), health facility data, translation
'names, and historic entries used across the linelist. Provides factory
'creation with worksheet validation, lazy-loaded translation caching via
'parallel BetterArray keys/values, hierarchical geo-level lookups using
'AutoFilter cascades, import/export in both flat-data and full-structure
'modes, and synchronisation of translated labels across linelist sheets,
'dictionary, and show/hide tables.
'@depends BetterArray, IPasswords, ILLdictionary, ICustomTable, ILLVariables, IDropdownLists, DropdownLists, LLVariables, CustomTable, HiddenNames, IHiddenNames

'@section State
'===============================================================================

Private Type TGeo
    geosh As Worksheet
    geoHidden As IHiddenNames
    namesCacheKeys As BetterArray
    namesCacheValues As BetterArray
    namesCacheLoaded As Boolean
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

'HiddenNames identifiers (worksheet-level, not cell-based ranges)
Private Const HN_GEO_UPDATED As String = "RNG_GeoUpdated"
Private Const HN_GEO_NAME As String = "RNG_GeoName"
Private Const HN_META_LANG As String = "RNG_MetaLang"
Private Const HN_GEO_LANG_CODE As String = "RNG_GeoLangCode"
Private Const HN_FORM_LOADED As String = "RNG_FormLoaded"

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
'@param sh Worksheet. The geobase worksheet containing T_ADM1 through T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF, T_METADATA, and required named ranges/hidden names.
'@return ILLGeo. A fully initialised LLGeo instance exposed through the ILLGeo interface.
'@export
Public Function Create(ByVal sh As Worksheet) As ILLGeo
    CheckRequirements sh

    With New LLGeo
        Set .Wksh = sh
        Set Create = .Self()
    End With
End Function

'@label:self
'@prop-title Self-reference for factory pattern.
'@details
'Returns a reference to the current instance cast as ILLGeo. Used inside the
'factory Create method to return the newly constructed object through its
'interface type.
'@return ILLGeo. The current instance.
Public Property Get Self() As ILLGeo
    Set Self = Me
End Property

'@label:wksh
'@prop-title Get the geobase worksheet.
'@details
'Returns the worksheet that holds all geobase ListObjects and named ranges.
'This worksheet is set during factory creation and should not change during
'the lifetime of the object.
'@return Worksheet. The geobase worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.geosh
End Property

'@label:wksh-set
'@prop-title Set the geobase worksheet.
'@details
'Assigns the internal worksheet reference. Called only from the factory Create
'method after validation. Not intended for external use.
'@param sh Worksheet. The geobase worksheet to bind to this instance.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.geosh = sh
End Property

'@section Hidden Names Store
'===============================================================================

'@label:geo-hidden-store
'@sub-title Lazy-loaded HiddenNames manager for the geo worksheet.
'@details
'Creates and caches an IHiddenNames instance bound to the geobase
'worksheet on first access. Used by all methods that read or write
'the 5 hidden name values (RNG_GeoUpdated, RNG_GeoName, RNG_MetaLang,
'RNG_GeoLangCode, RNG_FormLoaded).
'@return IHiddenNames. The worksheet-level hidden names manager.
Private Function GeoHiddenStore() As IHiddenNames
    If this.geoHidden Is Nothing Then
        Set this.geoHidden = HiddenNames.Create(Wksh())
    End If
    Set GeoHiddenStore = this.geoHidden
End Function

'@section Table Names
'===============================================================================

'@label:table-names
'@prop-title Ordered list of all geobase table names.
'@details
'Builds and returns a BetterArray containing the eight canonical ListObject
'names used on the geobase worksheet: T_ADM1, T_ADM2, T_ADM3, T_ADM4, T_HF,
'T_NAMES, T_HISTOGEO, and T_HISTOHF. The order matches the admin hierarchy
'followed by support tables.
'@return BetterArray. Ordered list of table name strings.
Public Property Get TableNames() As BetterArray
    Dim tabl As BetterArray
    Set tabl = New BetterArray
    tabl.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
              "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF
    Set TableNames = tabl.Clone()
End Property

'@section Translation Cache
'===============================================================================

'@label:ensure-names-cache
'@sub-title Ensure the translation name cache is loaded.
'@details
'Lazily initialises two parallel BetterArray instances (keys and values) that
'map raw admin column names (adm1_name, adm2_name, etc.) to their translated
'display names resolved from geobase named ranges or hidden names. If the
'cache is already loaded this method returns immediately. The worksheet is
'recalculated before reading to ensure formula-driven names are up to date.
Private Sub EnsureNamesCache()
    Dim sh As Worksheet

    If this.namesCacheLoaded Then Exit Sub

    Set sh = Wksh()
    sh.Calculate

    Set this.namesCacheKeys = New BetterArray
    Set this.namesCacheValues = New BetterArray

    CacheOneName sh, "adm1_name", "RNG_ADM1NAME"
    CacheOneName sh, "adm2_name", "RNG_ADM2NAME"
    CacheOneName sh, "adm3_name", "RNG_ADM3NAME"
    CacheOneName sh, "adm4_name", "RNG_ADM4NAME"
    CacheOneName sh, "hf_name", "RNG_HFNAME"
    CacheOneName sh, "geobase", HN_GEO_NAME

    this.namesCacheLoaded = True
End Sub

'@label:cache-one-name
'@sub-title Cache one translation mapping from raw name to resolved value.
'@details
'Reads a single value from either a HiddenName or a cell-based named range
'on the worksheet and stores the raw-name to resolved-value pair in the
'parallel cache arrays. Tries the HiddenNames store first; if not found
'falls back to sh.Range(). If neither exists or the value is not a non-empty
'string, the raw name is stored as both key and value so lookups fall back
'gracefully.
'@param sh Worksheet. The geobase worksheet to read from.
'@param rawName String. The raw column name used as the cache key.
'@param rngName String. The named range or hidden name identifier whose value provides the translation.
Private Sub CacheOneName(ByVal sh As Worksheet, ByVal rawName As String, _
                         ByVal rngName As String)
    Dim actualName As Variant
    Dim hn As IHiddenNames

    'Try HiddenNames first
    Set hn = GeoHiddenStore()
    If hn.HasName(rngName) Then
        actualName = hn.ValueAsString(rngName)
        If LenB(CStr(actualName)) > 0 Then
            this.namesCacheKeys.Push rawName
            this.namesCacheValues.Push CStr(actualName)
        Else
            this.namesCacheKeys.Push rawName
            this.namesCacheValues.Push rawName
        End If
        Exit Sub
    End If

    'Fall back to cell-based named range
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push rawName
        Exit Sub
    End If

    actualName = rng.Value
    If TypeName(actualName) = "String" Then
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push CStr(actualName)
    Else
        this.namesCacheKeys.Push rawName
        this.namesCacheValues.Push rawName
    End If
End Sub

'@label:reset-names-cache
'@sub-title Reset the translation cache.
'@details
'Marks the cache as unloaded and releases the key/value BetterArray references.
'Must be called after any operation that changes translation state (Import or
'Translate) so the next GeoNames access rebuilds the cache from fresh data.
Private Sub ResetNamesCache()
    this.namesCacheLoaded = False
    Set this.namesCacheKeys = Nothing
    Set this.namesCacheValues = Nothing
End Sub

'@label:geo-names
'@prop-title Translated display name for a raw admin column name.
'@details
'Converts a raw admin column name (such as "adm1_name", "hf_name", or
'"geobase") to its translated display name by searching the lazily loaded
'translation cache. If no mapping is found the raw name is returned unchanged.
'The cache is built on first access from geobase named ranges.
'@param admname String. Raw column name to translate.
'@return String. Translated display name, or the raw name if no translation exists.
'@export
Private Property Get GeoNames(ByVal admname As String) As String
    Dim counter As Long

    EnsureNamesCache

    For counter = this.namesCacheKeys.LowerBound To this.namesCacheKeys.UpperBound
        If this.namesCacheKeys.Item(counter) = admname Then
            GeoNames = CStr(this.namesCacheValues.Item(counter))
            Exit Property
        End If
    Next

    GeoNames = admname
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
'@param colName String. The column name to search for.
'@param headerRng Range. The header row range to search in.
'@param inRange Optional Boolean. True to return relative position, False for absolute column number. Defaults to True.
'@return Long. Column index (1-based), or -1 if not found.
Private Function ColumnIndex(ByVal colName As String, ByVal headerRng As Range, _
                             Optional ByVal inRange As Boolean = True) As Long
    Dim matchCell As Range

    Set matchCell = headerRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=True)

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
'If the geobase is in translated state the column name is first converted to
'its translated display name via GeoNames. Raises an error when the column
'cannot be found in the table header.
'@param colName String. The raw column name to look up.
'@param adminLevel Byte. Admin level index selecting which table to search (0-3 for T_ADM1 through T_ADM4).
'@return Long. 1-based column index within the table.
'@throws ProjectError.ElementNotFound When the column is not found in the admin table.
Private Function GeoColumnIndex(ByVal colName As String, ByVal adminLevel As Byte) As Long
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim loName As String
    Dim actualColName As String
    Dim colIdx As Long

    Set sh = Wksh()
    Set tabl = TableNames()
    loName = tabl.Item(adminLevel)
    actualColName = colName

    If GeoHiddenStore().ValueAsString(HN_GEO_UPDATED) = "updated, translated" Then
        actualColName = GeoNames(colName)
    End If

    colIdx = ColumnIndex(actualColName, sh.ListObjects(loName).HeaderRowRange)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & actualColName & " not found in table " & loName

    GeoColumnIndex = colIdx
End Function

'@label:hf-column-index
'@sub-title Resolve the column index for the HF table.
'@details
'Looks up a column name in the header of the T_HF ListObject. If the geobase
'is in translated state the column name is first converted to its translated
'display name via GeoNames. Raises an error when the column cannot be found.
'@param colName String. The raw column name to look up.
'@return Long. 1-based column index within the T_HF table.
'@throws ProjectError.ElementNotFound When the column is not found in the HF table.
Private Function HFColumnIndex(ByVal colName As String) As Long
    Dim sh As Worksheet
    Dim actualColName As String
    Dim colIdx As Long

    Set sh = Wksh()
    actualColName = colName

    If GeoHiddenStore().ValueAsString(HN_GEO_UPDATED) = "updated, translated" Then
        actualColName = GeoNames(colName)
    End If

    colIdx = ColumnIndex(actualColName, sh.ListObjects("T_" & HF).HeaderRowRange)
    If colIdx < 0 Then ThrowError ElementNotFound, _
        "Column " & actualColName & " not found in table T_HF"

    HFColumnIndex = colIdx
End Function

'@label:unique-range-value
'@sub-title Extract unique values from a range into a BetterArray.
'@details
'Iterates every cell in the supplied range and collects values that have not
'already been added, producing a deduplicated list in encounter order. Used
'internally to build dropdown-ready lists from filtered or unfiltered columns.
'@param rng Range. The range of cells to extract unique values from.
'@return BetterArray. Deduplicated values in encounter order.
Private Function UniqueRangeValue(ByVal rng As Range) As BetterArray
    Dim cellRng As Range
    Dim retVal As BetterArray

    Set retVal = New BetterArray

    For Each cellRng In rng
        If Not retVal.Includes(cellRng.Value) Then retVal.Push cellRng.Value
    Next

    Set UniqueRangeValue = retVal.Clone()
End Function

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

'@section GeoLevel
'===============================================================================

'@label:geo-level
'@prop-title List of values at a given admin level.
'@details
'Returns a BetterArray of unique geographic names at the requested admin level,
'optionally filtered by parent admin selections. For admin1 the full unfiltered
'list is returned. For admin2 through admin4 the caller supplies parent
'selections so the result cascades hierarchically via AutoFilter. The geoType
'parameter switches between the admin hierarchy and the health facility table.
'Calculation is temporarily disabled for performance during filtering.
'@param adminLevel Byte. Admin level index (0-3 for admin1-admin4).
'@param geoType Optional Byte. GeoScope value: GeoScopeAdmin (0) or GeoScopeHF (1). Defaults to GeoScopeAdmin.
'@param adminNames Optional Variant. Parent admin selections for filtering: String for level 2, BetterArray for levels 3-4. Defaults to vbNullString.
'@return BetterArray. Unique values at the requested level, or an empty array when the geobase is empty.
'@throws ProjectError.InvalidArgument When adminNames has wrong type or length.
'@throws ProjectError.ErrorUnexpectedState When the ListObject has no data.
'@export
Private Property Get GeoLevel(ByVal adminLevel As Byte, _
                              Optional ByVal geoType As Byte = GeoScopeAdmin, _
                              Optional adminNames As Variant = vbNullString) As BetterArray

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outLevels As BetterArray
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim rng As Range
    Dim loName As String
    Dim colIdx As Long
    Dim counter As Long
    Dim returnedRng As Range
    Dim prevCalc As Boolean

    Set sh = Wksh()
    Set tabl = TableNames()
    Set outLevels = New BetterArray

    'Disable calculation for performance during filtering
    prevCalc = sh.EnableCalculation
    sh.EnableCalculation = False

    On Error GoTo Cleanup

    If geoType = GeoScopeAdmin Then
        loName = tabl.Item(adminLevel)
    Else
        loName = tabl.Item(4) 'T_HF
    End If

    'Empty geobase: return empty array
    If GeoHiddenStore().ValueAsString(HN_GEO_UPDATED) = "empty" Then GoTo Cleanup

    'Validate adminNames
    If (adminLevel > 1) And (Not TypeName(adminNames) = "BetterArray") Then
        ThrowError InvalidArgument, _
            "Admin names should be a list of all previous admin levels for admin " & adminLevel
    ElseIf (adminLevel = LevelAdmin1 Or adminLevel = LevelAdmin2) _
            And (Not TypeName(adminNames) = "String") Then
        ThrowError InvalidArgument, _
            "No need to add admin names if you want the list of all admin"
    ElseIf TypeName(adminNames) = "BetterArray" Then
        If adminNames.Length <> CLng(adminLevel) Then
            ThrowError InvalidArgument, _
                "You should provide a table of length " & CLng(adminLevel) & _
                " for admin " & (CLng(adminLevel) + 1)
        End If
        If adminNames.LowerBound <> 1 Then
            ThrowError InvalidArgument, _
                "Please ensure your BetterArray has 1 as LowerBound, it is mandatory"
        End If
    End If

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range

    If Lo.DataBodyRange Is Nothing Then
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
            LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames
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
            LoRng.AutoFilter Field:=colIdx, Criteria1:=adminNames
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

    'For levels beyond admin1, extract filtered visible cells
    If adminLevel <> LevelAdmin1 Then
        On Error Resume Next
        Set returnedRng = Lo.ListColumns(colIdx).DataBodyRange _
                            .SpecialCells(xlCellTypeVisible)
        On Error GoTo Cleanup

        If Not returnedRng Is Nothing Then
            Lo.AutoFilter.ShowAllData
            Set outLevels = UniqueRangeValue(returnedRng)
        End If

        On Error Resume Next
        Lo.AutoFilter.ShowAllData
        On Error GoTo Cleanup
    End If

Cleanup:
    Set GeoLevel = outLevels.Clone()
    sh.EnableCalculation = prevCalc
End Property

'@section Translate
'===============================================================================

'@label:translate
'@sub-title Translate or revert geobase table headers.
'@details
'Applies or reverts translated display names on all geobase table headers
'(admin1 through admin4 and health facility). When rawNames is False, headers
'are translated to display names read from the geobase named ranges. When True,
'headers are reverted to raw column names. The method skips the operation when
'the geobase is already in the requested translation state and resets the
'translation cache both before and after the change.
'@param rawNames Optional Boolean. True to revert to raw names, False to apply translated names. Defaults to False.
'@export
Private Sub Translate(Optional ByVal rawNames As Boolean = False)
    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim counter As Long
    Dim counter2 As Long
    Dim Lo As ListObject
    Dim loName As String
    Dim admname As String
    Dim transGeoName As String
    Dim hRng As Range
    Dim cellRng As Range
    Dim geoUpdatedVal As String

    Set sh = Wksh()
    sh.Calculate
    geoUpdatedVal = GeoHiddenStore().ValueAsString(HN_GEO_UPDATED)

    'Skip if already in the requested state
    If Not rawNames And (InStr(1, geoUpdatedVal, ", translated") > 0) Then Exit Sub
    If rawNames And (InStr(1, geoUpdatedVal, ", not translated") > 0) Then Exit Sub

    'Reset cache so GeoNames reads fresh values
    ResetNamesCache

    Set tabl = TableNames()
    tabl.LowerBound = 1

    'Translate admin tables (T_ADM1 through T_ADM4)
    For counter = 1 To 4
        loName = tabl.Item(counter)
        Set Lo = sh.ListObjects(loName)
        Set hRng = Lo.HeaderRowRange

        For counter2 = 1 To counter
            admname = "adm" & counter2 & "_name"
            transGeoName = GeoNames(admname)
            hRng.Cells(1, counter2).Value = IIf(rawNames, admname, transGeoName)
        Next
    Next

    'Translate HF table headers
    Set Lo = sh.ListObjects("T_HF")
    Set hRng = Lo.HeaderRowRange
    transGeoName = GeoNames("hf_name")
    hRng.Cells(1, 1).Value = IIf(rawNames, "hf_name", transGeoName)
    Set cellRng = hRng.Cells(1, 5)

    For counter = 1 To 3
        admname = "adm" & counter & "_name"
        transGeoName = GeoNames(admname)
        cellRng.Offset(, -counter).Value = IIf(rawNames, admname, transGeoName)
    Next

    'Update translation state flag via HiddenNames
    geoUpdatedVal = Replace(geoUpdatedVal, ", not translated", vbNullString)
    geoUpdatedVal = Replace(geoUpdatedVal, ", translated", vbNullString)
    geoUpdatedVal = IIf(rawNames, geoUpdatedVal & ", not translated", geoUpdatedVal & ", translated")
    GeoHiddenStore().SetValue HN_GEO_UPDATED, geoUpdatedVal

    'Reset cache after translation change
    ResetNamesCache
End Sub

'@section Clear
'===============================================================================

'@label:clear
'@sub-title Clear all geobase data.
'@details
'Removes all data from every ListObject on the geobase worksheet except
'T_LevelNames and T_CONCAT, and sets the update flag to "empty". Resets the
'translation name cache so subsequent GeoNames lookups will rebuild from
'fresh data.
'@export
Private Sub Clear()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = Wksh()

    For Each Lo In sh.ListObjects
        On Error Resume Next
        If Lo.Name <> "T_LevelNames" And Lo.Name <> "T_CONCAT" Then
            Lo.DataBodyRange.Cells.Clear
        End If
        On Error GoTo 0
    Next

    GeoHiddenStore().SetValue HN_GEO_UPDATED, "empty"
    ResetNamesCache
End Sub

'@label:clear-historic
'@sub-title Clear historic entries.
'@details
'Clears the data body of historic ListObjects. The scope parameter controls
'which tables are cleared: 0 for geo only (T_HISTOGEO), 1 for HF only
'(T_HISTOHF), and 3 for both. Errors are silently suppressed so the method
'succeeds even when a table has no data body.
'@param scope Optional Byte. 0 = geo only, 1 = HF only, 3 = both. Defaults to 3.
'@export
Private Sub ClearHistoric(Optional ByVal scope As Byte = 3)
    Dim sh As Worksheet
    Set sh = Wksh()

    On Error Resume Next
    If (scope = 0) Or (scope = 3) Then
        sh.ListObjects("T_" & HISTOGEO).DataBodyRange.Cells.Clear
    End If
    If (scope = 1) Or (scope = 3) Then
        sh.ListObjects("T_" & HISTOHF).DataBodyRange.Cells.Clear
    End If
    On Error GoTo 0
End Sub

'@section Import
'===============================================================================

'@label:import
'@sub-title Import geobase data from an external workbook.
'@details
'Imports geographic data from the supplied workbook into the geobase worksheet.
'Each worksheet whose name matches a known table name (ADM1 through ADM4, HF,
'NAMES, HISTOGEO, HISTOHF, METADATA) is imported column by column into the
'corresponding ListObject. Concat formulas are regenerated for filtered lookup
'support. When histoOnly is True only historic tables are refreshed and the
'main admin/HF data is left intact. After import, the geobase is marked as
'updated and headers are translated.
'@param fromWkb Workbook. Source workbook containing geo data sheets.
'@param histoOnly Optional Boolean. True to import only historic data. Defaults to False.
'@export
Private Sub Import(ByVal fromWkb As Workbook, _
                   Optional ByVal histoOnly As Boolean = False)

    Dim sh As Worksheet
    Dim impsh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim admNames As BetterArray
    Dim counter As Long
    Dim colName As String
    Dim colIdx As Long
    Dim lastRow As Long
    Dim impRng As Range
    Dim actRng As Range
    Dim concatFormula As String
    Dim sep As String
    Dim concatRng As Range

    Set admNames = New BetterArray
    admNames.Push HISTOGEO, HISTOHF, METADATA

    If histoOnly Then
        ClearHistoric
    Else
        Clear
        admNames.Push ADMIN1, ADMIN2, ADMIN3, ADMIN4, HF, NAMES
    End If

    Set sh = Wksh()
    sep = " & " & Chr(34) & " | " & Chr(34) & " & "

    Translate rawNames:=True

    For Each impsh In fromWkb.Worksheets
        With impsh
            If admNames.Includes(.Name) Then
                lastRow = Application.WorksheetFunction.Max(.UsedRange.Rows.Count, 2)
                Set Lo = sh.ListObjects("T_" & impsh.Name)
                Set hRng = Lo.HeaderRowRange

                'Import data column by column
                counter = 1
                Do While impsh.Cells(1, counter) <> vbNullString
                    colName = .Cells(1, counter).Value
                    colIdx = ColumnIndex(colName, hRng, inRange:=False)
                    If colIdx > 0 Then
                        Set impRng = .Range(.Cells(2, counter), .Cells(lastRow, counter))
                        Set actRng = sh.Range(sh.Cells(2, colIdx), sh.Cells(lastRow, colIdx))
                        actRng.Value = impRng.Value
                    End If
                    counter = counter + 1
                Loop

                'Build concat formula for the last column of the table
                Set impRng = Lo.HeaderRowRange.Offset(1)

                Select Case Lo.Name
                Case "T_ADM1"
                    concatFormula = "=" & impRng.Cells(1, 1).Address(RowAbsolute:=False)

                Case "T_ADM2"
                    concatFormula = "=" & impRng.Cells(1, 1).Address(RowAbsolute:=False) & _
                                   sep & impRng.Cells(1, 2).Address(RowAbsolute:=False)

                Case "T_ADM3"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False)

                Case "T_ADM4"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 4).Address(RowAbsolute:=False)

                Case "T_HF"
                    concatFormula = "=" & _
                        impRng.Cells(1, 1).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 2).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 3).Address(RowAbsolute:=False) & sep & _
                        impRng.Cells(1, 4).Address(RowAbsolute:=False)

                Case Else
                    concatFormula = vbNullString
                End Select

                'Clear concat column before resizing
                Set concatRng = sh.Range( _
                    sh.Cells(2, hRng.Column + hRng.Cells.Count - 1), _
                    sh.Cells(Lo.Range.Rows.Count, hRng.Column + hRng.Cells.Count - 1))

                If concatFormula <> vbNullString Then concatRng.ClearContents

                'Resize the ListObject to fit imported data
                Set actRng = sh.Range( _
                    sh.Cells(1, hRng.Column), _
                    sh.Cells(lastRow, hRng.Column + hRng.Cells.Count - 1))
                Lo.Resize actRng

                'Apply the concat formula
                If concatFormula <> vbNullString Then
                    sh.Cells(impRng.Row, hRng.Column + hRng.Cells.Count - 1).Formula = _
                        concatFormula
                End If
            End If
        End With
    Next

    GeoHiddenStore().SetValue HN_GEO_NAME, Replace(fromWkb.Name, ".xlsx", vbNullString)
    GeoHiddenStore().SetValue HN_GEO_UPDATED, "updated"
    Translate rawNames:=False

    UpdateMetadata
    ResetNamesCache
End Sub

'@section Export
'===============================================================================

'@label:export-to-file
'@sub-title Export geobase tables to a workbook as flat data.
'@details
'Exports geobase tables to the target workbook with one sheet per table and a
'Metadata sheet. Headers are reverted to raw names before export and restored
'to translated form afterwards. When onlyHistory is True only the HISTOGEO and
'HISTOHF tables are exported.
'@param outputWkb Workbook. Target workbook to receive exported sheets.
'@param onlyHistory Optional Boolean. True to export only historic tables. Defaults to False.
'@export
Private Sub ExportToFile(ByVal outputWkb As Workbook, _
                         Optional ByVal onlyHistory As Boolean = False)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim tabl As BetterArray
    Dim outputRng As Range
    Dim inputRng As Range
    Dim sheetName As String
    Dim loName As String
    Dim counter As Long

    Set tabl = New BetterArray
    Set sh = Wksh()
    Translate rawNames:=True

    If onlyHistory Then
        tabl.Push "T_" & HISTOGEO, "T_" & HISTOHF
    Else
        Set tabl = TableNames()
    End If

    With outputWkb
        'Export metadata sheet first
        .Worksheets(1).Name = METADATASHEET
        Set inputRng = sh.ListObjects("T_" & METADATA).Range

        With .Worksheets(METADATASHEET)
            Set outputRng = .Range(.Cells(1, 1), _
                .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
        End With
        outputRng.Value = inputRng.Value
        sheetName = METADATASHEET

        'Export each table to its own sheet
        For counter = tabl.LowerBound To tabl.UpperBound
            loName = tabl.Item(counter)
            Set inputRng = sh.ListObjects(loName).Range
            .Worksheets.Add(After:=.Worksheets(sheetName)).Name = _
                Replace(loName, "T_", vbNullString)
            sheetName = Replace(loName, "T_", vbNullString)

            With .Worksheets(sheetName)
                Set outputRng = .Range(.Cells(1, 1), _
                    .Cells(inputRng.Rows.Count, inputRng.Columns.Count))
            End With
            outputRng.Value = inputRng.Value
        Next
    End With

    Translate rawNames:=False
End Sub

'@label:export-to-wkb
'@sub-title Export the geobase worksheet structure to a workbook.
'@details
'Copies the entire geobase worksheet to the target workbook, preserving
'ListObjects, named ranges, formulas, pcode references, concat references,
'and metadata. Creates a hidden Metadata sheet with creation metadata including
'linelist name, setup name, language, and timestamp. Cell-based named ranges
'(RNG_PastingGeoCol, RNG_HFNAME, RNG_ADM1NAME-4) are exported as regular
'named ranges. HiddenNames entries (RNG_GeoUpdated, RNG_GeoName, etc.) are
'exported via the HiddenNames manager. Raises an error when sheets with
'conflicting names already exist in the target workbook.
'@param outputWkb Workbook. Target workbook.
'@param llName Optional String. Linelist name for metadata. Defaults to vbNullString.
'@param setupPath Optional String. Setup file path for metadata. Defaults to vbNullString.
'@throws ProjectError.ElementShouldNotExists When target sheets already exist.
'@export
Private Sub ExportToWkb(ByVal outputWkb As Workbook, _
                        Optional ByVal llName As String = vbNullString, _
                        Optional ByVal setupPath As String = vbNullString)

    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As ListObject
    Dim expAddress As String
    Dim expRange As Range
    Dim rng As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String
    Dim metaKeys As BetterArray
    Dim metaValues As BetterArray
    Dim setupName As String

    Set sh = Wksh()

    'Check that target sheets do not already exist
    On Error Resume Next
    Set outsh = outputWkb.Worksheets(sh.Name)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & sh.Name & " in the workbook"
    End If

    On Error Resume Next
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    On Error GoTo 0
    If Not outsh Is Nothing Then
        ThrowError ElementShouldNotExists, _
            "Unable to export, already a sheet named " & METADATASHEET & " in the workbook"
    End If

    Translate rawNames:=False

    'Create the geo sheet and copy ListObjects
    outputWkb.Worksheets.Add.Name = sh.Name
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

    outsh.Range("A:AZ").Locked = False
    outsh.Range("A:AZ").EntireColumn.AutoFit

    'Export cell-based named ranges (6 ranges that remain as cells)
    Set rngTable = New BetterArray
    rngTable.Push "RNG_PastingGeoCol", "RNG_HFNAME", "RNG_ADM1NAME", _
                  "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME"

    For counter = rngTable.LowerBound To rngTable.UpperBound
        rngName = rngTable.Item(counter)
        expAddress = sh.Range(rngName).Address
        Set expRange = outsh.Range(expAddress)

        expRange.Name = rngName
        expRange.Value = sh.Range(rngName).Value
        expRange.Locked = True

        If sh.Range(rngName).Formula <> vbNullString Then
            On Error Resume Next
            expRange.Formula = sh.Range(rngName).Formula
            On Error GoTo 0
        End If
    Next

    'Export HiddenNames entries (5 ranges stored as hidden names)
    GeoHiddenStore().ExportNames outsh

    'Create metadata sheet
    outputWkb.Worksheets.Add.Name = METADATASHEET
    Set outsh = outputWkb.Worksheets(METADATASHEET)
    Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)

    On Error Resume Next
    setupName = Split(setupPath, Application.PathSeparator)( _
        UBound(Split(setupPath, Application.PathSeparator)))
    On Error GoTo 0

    Set metaKeys = New BetterArray
    Set metaValues = New BetterArray

    metaKeys.Push "variable", "language", "linelist_creation_date", _
                  "linelist_creation_name", "used_setup_name", "used_designer_name"

    metaValues.Push "value", GeoHiddenStore().ValueAsString(HN_META_LANG), _
                    Format(Now(), "yyyy-mm-dd Hh:Nn"), _
                    llName, setupName, ThisWorkbook.Name

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
End Sub

'@section Update
'===============================================================================

'@label:update
'@sub-title Update linelist geo column headers and dropdowns.
'@details
'Iterates all HList worksheets in the parent workbook, finds columns controlled
'by geo1 through geo4 and hf, and updates their header labels with translated
'geo display names. Also updates corresponding labels on print sheets. Finally,
'refreshes the admin1 dropdown list on the dropdown_lists__ sheet so that it
'reflects the current geobase content.
'@param pass IPasswords. Password manager for sheet protection and unprotection.
'@export
Private Sub Update(ByVal pass As IPasswords)

    Const DROPDOWNSHEET As String = "dropdown_lists__"

    Dim sh As Worksheet
    Dim wb As Workbook
    Dim cellRng As Range
    Dim controlValue As String
    Dim geoData As BetterArray
    Dim drop As IDropdownLists
    Dim cellPrintRng As Range
    Dim printsh As Worksheet

    Set wb = Wksh.Parent

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            pass.UnProtect sh.Name

            With sh
                Set cellRng = .Range(.Cells(1, 4).Value & "_START")
                Set printsh = wb.Worksheets("print_" & sh.Name)
                Set cellPrintRng = printsh.Range(cellRng.Address)

                Do While Not IsEmpty(cellRng)
                    controlValue = cellRng.Offset(-4).Value

                    Select Case controlValue
                    Case "geo1"
                        cellRng.Offset(-1).Value = GeoNames("adm1_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm1_name")
                    Case "geo2"
                        cellRng.Offset(-1).Value = GeoNames("adm2_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm2_name")
                    Case "geo3"
                        cellRng.Offset(-1).Value = GeoNames("adm3_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm3_name")
                    Case "geo4"
                        cellRng.Offset(-1).Value = GeoNames("adm4_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("adm4_name")
                    Case "hf"
                        cellRng.Offset(-1).Value = GeoNames("hf_name")
                        cellPrintRng.Offset(-1).Value = GeoNames("hf_name")
                    End Select

                    Set cellRng = cellRng.Offset(, 1)
                    Set cellPrintRng = cellPrintRng.Offset(, 1)
                Loop
            End With

            pass.Protect sh.Name
        End If
    Next

    'Update the admin1 dropdown
    On Error Resume Next
    Set sh = wb.Worksheets(DROPDOWNSHEET)
    On Error GoTo 0

    If sh Is Nothing Then Exit Sub

    Set drop = DropdownLists.Create(sh)
    Set geoData = GeoLevel(LevelAdmin1)
    drop.Update geoData, "admin1"
End Sub

'@label:update-dict
'@sub-title Update dictionary labels for geo-related variables.
'@details
'Scans the dictionary control column for geo1 through geo4 and hf entries and
'overwrites the corresponding main label cells with translated geo display
'names from the geobase. This keeps the dictionary in sync after a geobase
'import or translation change.
'@param dict ILLdictionary. The linelist dictionary whose labels should be updated.
'@export
Private Sub UpdateDict(ByVal dict As ILLdictionary)
    Dim controlRng As Range
    Dim mainlabRng As Range
    Dim counter As Long

    Set controlRng = dict.DataRange("control")
    Set mainlabRng = dict.DataRange("main label")

    For counter = 1 To controlRng.Rows.Count
        Select Case controlRng.Cells(counter, 1).Value
        Case "geo1"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm1_name")
        Case "geo2"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm2_name")
        Case "geo3"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm3_name")
        Case "geo4"
            mainlabRng.Cells(counter, 1).Value = GeoNames("adm4_name")
        Case "hf"
            mainlabRng.Cells(counter, 1).Value = GeoNames("hf_name")
        End Select
    Next
End Sub

'@label:update-show-hide
'@sub-title Update show/hide tables with translated geo labels.
'@details
'Iterates all ListObjects on the show/hide worksheet, finds variables
'controlled by geo1 through geo4 and hf via the dictionary, and updates their
'main label cells to include the subsection prefix followed by the translated
'geo display name. This keeps the show/hide interface labels in sync with the
'current geobase translation.
'@param sh Worksheet. The show/hide worksheet containing ListObjects to update.
'@param dict ILLdictionary. The linelist dictionary for resolving control and subsection values.
'@export
Private Sub UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
    Dim csTab As ICustomTable
    Dim vars As ILLVariables
    Dim Lo As ListObject
    Dim counter As Long
    Dim mainlabRng As Range
    Dim varRng As Range
    Dim controlValue As String
    Dim varValue As String
    Dim subsecValue As String

    Set vars = LLVariables.Create(dict)

    For Each Lo In sh.ListObjects
        Set csTab = CustomTable.Create(Lo)
        Set mainlabRng = csTab.DataRange("main label")
        Set varRng = csTab.DataRange("variable name")

        For counter = 1 To varRng.Rows.Count
            varValue = varRng.Cells(counter, 1).Value
            controlValue = vars.Value(colName:="control", varName:=varValue)
            subsecValue = vars.Value(colName:="sub section", varName:=varValue)

            Select Case controlValue
            Case "geo1"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm1_name")
            Case "geo2"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm2_name")
            Case "geo3"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm3_name")
            Case "geo4"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("adm4_name")
            Case "hf"
                mainlabRng.Cells(counter, 1).Value = _
                    subsecValue & ": " & GeoNames("hf_name")
            End Select
        Next
    Next
End Sub

'@section Historic
'===============================================================================

'@label:update-historic
'@sub-title Add a new entry to the historic data.
'@details
'Appends a value to the appropriate historic ListObject (T_HISTOGEO or
'T_HISTOHF), resizes the table to include the new row, removes duplicates,
'and sorts alphabetically. Handles the edge case where the ListObject has no
'data body or is empty by inserting at the first data row.
'@param newValue String. The value to add to the historic list.
'@param scope Optional Byte. 0 = geo historic (T_HISTOGEO), 1 = HF historic (T_HISTOHF). Defaults to 0.
'@export
Private Sub UpdateHistoric(ByVal newValue As String, _
                           Optional ByVal scope As Byte = 0)
    Dim sh As Worksheet
    Dim loName As String
    Dim LoRng As Range
    Dim cellRng As Range
    Dim Lo As ListObject

    Set sh = Wksh()
    loName = IIf(scope = 0, HISTOGEO, HISTOHF)
    loName = "T_" & loName

    Set Lo = sh.ListObjects(loName)
    Set LoRng = Lo.Range
    Set cellRng = LoRng.Cells(LoRng.Rows.Count + 1, 1)

    If Lo.DataBodyRange Is Nothing Then
        Set cellRng = LoRng.Cells(2, 1)
    ElseIf IsEmpty(Lo.DataBodyRange) Then
        Set cellRng = LoRng.Cells(2, 1)
    End If

    cellRng.Value = newValue
    Lo.Resize sh.Range(LoRng.Cells(1, 1), cellRng)

    On Error Resume Next
    LoRng.RemoveDuplicates Columns:=1, Header:=xlYes
    LoRng.Sort key1:=LoRng, Header:=xlYes
    On Error GoTo 0
End Sub

'@section AdminCode
'===============================================================================

'@label:admin-code
'@sub-title Resolve the admin code from a translated display name.
'@details
'Compares the supplied translated admin name against GeoNames for each admin
'level and returns the matching raw code ("adm1" through "adm4"). Falls back
'to "adm1" when no match is found, which serves as a safe default for
'downstream callers.
'@param selectedAdmin String. The translated admin display name to resolve.
'@return String. Admin code ("adm1", "adm2", "adm3", or "adm4").
'@export
Private Function AdminCode(ByVal selectedAdmin As String) As String
    Select Case selectedAdmin
    Case GeoNames("adm1_name")
        AdminCode = "adm1"
    Case GeoNames("adm2_name")
        AdminCode = "adm2"
    Case GeoNames("adm3_name")
        AdminCode = "adm3"
    Case GeoNames("adm4_name")
        AdminCode = "adm4"
    Case Else
        AdminCode = "adm1"
    End Select
End Function

'@label:has-no-data
'@prop-title Whether the geobase has no imported data.
'@details
'Returns True when the geobase update flag (RNG_GeoUpdated) reads "empty",
'indicating that no geo data has been imported yet or that all data has been
'cleared. Used by consumers to skip geo-dependent operations when no data is
'available.
'@return Boolean. True when the geobase is empty.
'@export
Private Property Get HasNoData() As Boolean
    HasNoData = (GeoHiddenStore().ValueAsString(HN_GEO_UPDATED) = "empty")
End Property

'@section Internal Helpers
'===============================================================================

'@label:update-metadata
'@sub-title Update the metadata sheet from the geo metadata table.
'@details
'Copies the current content of T_METADATA into the hidden Metadata worksheet
'at the RNG_Meta_Anchor position. Clears any previously written metadata rows
'before writing the new data. Silently exits if the Metadata sheet or anchor
'range does not exist.
Private Sub UpdateMetadata()
    Const METADATASHEET As String = "Metadata"

    Dim sh As Worksheet
    Dim metash As Worksheet
    Dim rng As Range
    Dim expRng As Range
    Dim lastRow As Long

    On Error GoTo EndUpdate

    Set sh = Wksh()
    Set metash = sh.Parent.Worksheets(METADATASHEET)

    With metash
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set expRng = .Range("RNG_Meta_Anchor")

        If lastRow >= expRng.Row Then
            Set rng = .Range(expRng, .Cells(lastRow, 2))
            rng.ClearContents
        End If

        Set rng = sh.ListObjects("T_" & METADATA).Range.Offset(1)
        Set expRng = .Range(expRng, expRng.Offset(rng.Rows.Count - 1, 1))
        expRng.Value = rng.Value
    End With

EndUpdate:
End Sub

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate that the worksheet has all required tables, named ranges, and hidden names.
'@details
'Checks that the supplied worksheet is not Nothing, then verifies the presence
'of all nine required ListObjects (T_ADM1 through T_ADM4, T_HF, T_NAMES,
'T_HISTOGEO, T_HISTOHF, T_METADATA), five required HiddenNames entries
'(RNG_GeoUpdated, RNG_GeoName, RNG_GeoLangCode, RNG_FormLoaded, RNG_MetaLang),
'and six required cell-based named ranges (RNG_PastingGeoCol, RNG_HFNAME,
'RNG_ADM1NAME through RNG_ADM4NAME). Raises an appropriate error when any
'requirement is missing.
'@param sh Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.ElementNotFound When required tables, named ranges, or hidden names are missing.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim admTableNames As BetterArray
    Dim geoStore As IHiddenNames

    If sh Is Nothing Then
        ThrowError ObjectNotInitialized, "Geo sheet not defined"
    End If

    Set admTableNames = New BetterArray
    admTableNames.Push "T_" & ADMIN1, "T_" & ADMIN2, "T_" & ADMIN3, "T_" & ADMIN4, _
                       "T_" & HF, "T_" & NAMES, "T_" & HISTOGEO, "T_" & HISTOHF, _
                       "T_" & METADATA
    LoExists admTableNames, sh

    'Validate HiddenNames entries (5 names stored as hidden names)
    Set geoStore = HiddenNames.Create(sh)
    HiddenNameRequired geoStore, HN_GEO_NAME
    HiddenNameRequired geoStore, HN_GEO_UPDATED
    HiddenNameRequired geoStore, HN_GEO_LANG_CODE
    HiddenNameRequired geoStore, HN_FORM_LOADED
    HiddenNameRequired geoStore, HN_META_LANG

    'Validate cell-based named ranges (6 ranges that remain as cells)
    RangeExists sh, "RNG_PastingGeoCol"
    RangeExists sh, "RNG_HFNAME"
    RangeExists sh, "RNG_ADM1NAME"
    RangeExists sh, "RNG_ADM2NAME"
    RangeExists sh, "RNG_ADM3NAME"
    RangeExists sh, "RNG_ADM4NAME"
End Sub

'@label:hidden-name-required
'@sub-title Verify that a hidden name exists in the store.
'@details
'Checks whether the specified name exists in the given HiddenNames store.
'If not found, raises an ElementNotFound error with a descriptive message.
'@param store IHiddenNames. The hidden names manager to check.
'@param nameId String. The hidden name identifier to look for.
'@throws ProjectError.ElementNotFound When the hidden name does not exist.
Private Sub HiddenNameRequired(ByVal store As IHiddenNames, ByVal nameId As String)
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

'@section Interface Implementation
'===============================================================================

Private Property Get ILLGeo_GeoLevel(ByVal adminLevel As Byte, _
                                     Optional ByVal geoType As Byte = GeoScopeAdmin, _
                                     Optional adminNames As Variant = vbNullString _
                                     ) As BetterArray
    Set ILLGeo_GeoLevel = GeoLevel(adminLevel, geoType, adminNames)
End Property

Private Property Get ILLGeo_GeoNames(ByVal admname As String) As String
    ILLGeo_GeoNames = GeoNames(admname)
End Property

Private Property Get ILLGeo_HasNoData() As Boolean
    ILLGeo_HasNoData = HasNoData()
End Property

Private Sub ILLGeo_Translate(Optional ByVal rawNames As Boolean = False)
    Translate rawNames
End Sub

Private Sub ILLGeo_Import(ByVal fromWkb As Workbook, _
                          Optional ByVal histoOnly As Boolean = False)
    Import fromWkb, histoOnly
End Sub

Private Sub ILLGeo_ExportToFile(ByVal outputWkb As Workbook, _
                                Optional ByVal onlyHistory As Boolean = False)
    ExportToFile outputWkb, onlyHistory
End Sub

Private Sub ILLGeo_ExportToWkb(ByVal Wkb As Workbook, _
                               Optional ByVal llName As String = vbNullString, _
                               Optional ByVal setupPath As String = vbNullString)
    ExportToWkb Wkb, llName, setupPath
End Sub

Private Sub ILLGeo_Clear()
    Clear
End Sub

Private Sub ILLGeo_ClearHistoric(Optional ByVal scope As Byte = 3)
    ClearHistoric scope
End Sub

Private Sub ILLGeo_UpdateHistoric(ByVal newValue As String, _
                                  Optional ByVal scope As Byte = 0)
    UpdateHistoric newValue, scope
End Sub

Private Sub ILLGeo_Update(ByVal pass As IPasswords)
    Update pass
End Sub

Private Sub ILLGeo_UpdateDict(ByVal dict As ILLdictionary)
    UpdateDict dict
End Sub

Private Sub ILLGeo_UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
    UpdateShowHide sh, dict
End Sub

Private Function ILLGeo_AdminCode(ByVal selectedAdmin As String) As String
    ILLGeo_AdminCode = AdminCode(selectedAdmin)
End Function
