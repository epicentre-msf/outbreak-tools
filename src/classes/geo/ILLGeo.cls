VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for geobase management"

'@Interface
'@Folder("Geo")
'@ModuleDescription("Interface for geobase management")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILLGeo
'@description
'Defines the public contract for the LLGeo class, exposing methods for
'importing, exporting, translating, and querying the geobase data that
'underpins geographic dropdowns and spatial analyses in the linelist.
'Consumers use this interface to manage the four-level administrative
'hierarchy (admin1 through admin4) plus health facility data, translate
'column headers between raw and display names, maintain historic entries,
'and synchronise geo labels across linelist worksheets and the dictionary.

'@details
'The geobase is stored on a dedicated worksheet containing ListObjects
'for each admin level (T_ADM1 through T_ADM4), health facilities (T_HF),
'translation names (T_NAMES), historic entries (T_HISTOGEO, T_HISTOHF),
'and metadata (T_METADATA). Three enums control scope selection.
Public Enum GeoScope
    GeoScopeAdmin = 0
    GeoScopeHF = 1
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

'@details
'Extended scope that includes both admin levels and health facility. Used
'for operations that need to address all five geographic tiers.
Public Enum AdminAndHFScope
    LevelGeoAdmin1 = 0
    LevelGeoAdmin2 = 1
    LevelGeoAdmin3 = 2
    LevelGeoAdmin4 = 3
    LevelHF = 4
End Enum


'@section Properties
'===============================================================================

'@jump:geo-level
'@prop-title List of values at a given admin level
'@details
'Returns a BetterArray of unique geographic names at the requested admin
'level, optionally filtered by parent admin selections. For admin1 the
'full unfiltered list is returned. For admin2 through admin4, the caller
'supplies parent selections so the result cascades hierarchically. The
'geoType parameter switches between the admin hierarchy and the health
'facility table.
'@param adminLevel Byte. Admin level index (0-3 for admin1-admin4).
'@param geoType Optional Byte. GeoScope value: GeoScopeAdmin (0) or GeoScopeHF (1). Defaults to GeoScopeAdmin.
'@param adminNames Optional Variant. Parent admin selections for filtering: String for level 2, BetterArray for levels 3-4. Defaults to vbNullString.
'@return BetterArray. Unique values at the requested level, or an empty array when the geobase is empty.
'@export
Public Property Get GeoLevel(ByVal adminLevel As Byte, _
                             Optional ByVal geoType As Byte = GeoScopeAdmin, _
                             Optional adminNames As Variant = vbNullString) As BetterArray
End Property

'@jump:geo-names
'@prop-title Translated display name for a raw admin column name
'@details
'Converts a raw admin column name (such as "adm1_name", "hf_name", or
'"geobase") to its translated display name as configured in the geobase
'named ranges. Returns the raw name unchanged when no translation mapping
'exists.
'@param admname String. Raw column name to translate.
'@return String. Translated display name, or the raw name if no translation exists.
'@export
Public Property Get GeoNames(ByVal admname As String) As String: End Property

'@jump:has-no-data
'@prop-title Whether the geobase has no imported data
'@details
'Returns True when the geobase update flag reads "empty", indicating that
'no geo data has been imported yet or that all data has been cleared.
'@return Boolean. True when the geobase is empty.
'@export
Public Property Get HasNoData() As Boolean: End Property

Public Property Get Wksh() As Worksheet : End Property


'@section Translation
'===============================================================================

'@jump:translate
'@sub-title Translate or revert geobase table headers
'@details
'Applies or reverts translated display names on all geobase table headers
'(admin1 through admin4, health facility). When rawNames is False, headers
'are translated to display names. When True, headers are reverted to raw
'column names. Skips the operation when the geobase is already in the
'requested translation state.
'@param rawNames Optional Boolean. True to revert to raw names, False to apply translated names. Defaults to False.
'@export
Public Sub Translate(Optional ByVal rawNames As Boolean = False): End Sub


'@section Import / Export
'===============================================================================

'@jump:import
'@sub-title Import geobase data from an external workbook
'@details
'Imports geographic data from the supplied workbook. Each worksheet whose
'name matches a known table name (ADM1 through ADM4, HF, NAMES, HISTOGEO,
'HISTOHF, METADATA) is imported column by column into the corresponding
'ListObject. When histoOnly is True, only historic tables are refreshed
'and the main admin/HF data is left intact.
'@param fromWkb Workbook. Source workbook containing geo data sheets.
'@param histoOnly Optional Boolean. True to import only historic data. Defaults to False.
'@export
Public Sub Import(ByVal fromWkb As Workbook, _
                  Optional ByVal histoOnly As Boolean = False)
End Sub

'@jump:export-to-file
'@sub-title Export geobase tables to a workbook as flat data
'@details
'Exports geobase tables to the target workbook with one sheet per table
'and a Metadata sheet. Headers are reverted to raw names before export
'and restored to translated form afterwards. HiddenNames entries are
'exported to the Metadata sheet for round-trip import support. When
'onlyHistory is True, only the HISTOGEO and HISTOHF tables are exported.
'@param outputWkb Workbook. Target workbook to receive exported sheets.
'@param onlyHistory Optional Boolean. True to export only historic tables. Defaults to False.
'@export
Public Sub ExportToFile(ByVal outputWkb As Workbook, _
                        Optional ByVal onlyHistory As Boolean = False)
End Sub

'@jump:export-to-wkb
'@sub-title Export the geobase worksheet structure to a workbook
'@details
'Copies the entire geobase worksheet to the target workbook, preserving
'ListObjects, named ranges, formulas, pcode references, concat references,
'and metadata. Creates a hidden Metadata sheet with creation metadata.
'Raises an error when sheets with conflicting names already exist.
'@param Wkb Workbook. Target workbook.
'@param llName Optional String. Linelist name for metadata. Defaults to vbNullString.
'@param setupPath Optional String. Setup file path for metadata. Defaults to vbNullString.
'@export
Public Sub ExportToWkb(ByVal Wkb As Workbook, _
                       Optional ByVal llName As String = vbNullString, _
                       Optional ByVal setupPath As String = vbNullString)
End Sub


'@section Clear
'===============================================================================

'@jump:clear
'@sub-title Clear all geobase data
'@details
'Removes all data from every ListObject on the geobase worksheet except
'T_LevelNames and T_CONCAT, and sets the update flag to "empty". Resets
'the translation name cache.
'@export
Public Sub Clear(): End Sub

'@jump:clear-historic
'@sub-title Clear historic entries
'@details
'Clears the data body of historic ListObjects. The scope parameter
'controls which tables are cleared: 0 for geo only, 1 for HF only,
'3 for both.
'@param scope Optional Byte. 0 = geo only, 1 = HF only, 3 = both. Defaults to 3.
'@export
Public Sub ClearHistoric(Optional ByVal scope As Byte = 3): End Sub

'@jump:update-historic
'@sub-title Add a new entry to the historic data
'@details
'Appends a value to the appropriate historic ListObject (HISTOGEO or
'HISTOHF), resizes the table to include the new row, removes duplicates,
'and sorts alphabetically.
'@param newValue String. The value to add.
'@param scope Optional Byte. 0 = geo historic, 1 = HF historic. Defaults to 0.
'@export
Public Sub UpdateHistoric(ByVal newValue As String, _
                          Optional ByVal scope As Byte = 0)
End Sub


'@section Update
'===============================================================================

'@jump:update
'@sub-title Update linelist geo column headers and dropdowns
'@details
'Iterates all HList worksheets, finds columns controlled by geo1 through
'geo4 and hf, and updates their header labels with translated geo names.
'Also refreshes the admin1 dropdown list on the dropdown_lists__ sheet.
'@param pass IPasswords. Password manager for sheet protection and unprotection.
'@export
Public Sub Update(ByVal pass As IPasswords): End Sub

'@jump:update-dict
'@sub-title Update dictionary labels for geo-related variables
'@details
'Scans the dictionary control column for geo1 through geo4 and hf entries
'and overwrites the corresponding main label cells with translated geo
'display names.
'@param dict ILLdictionary. The linelist dictionary.
'@export
Public Sub UpdateDict(ByVal dict As ILLdictionary): End Sub

'@jump:update-show-hide
'@sub-title Update show/hide tables with translated geo labels
'@details
'Iterates all ListObjects on the show/hide worksheet, finds variables
'controlled by geo1 through geo4 and hf, and updates their main label
'cells to include the subsection prefix and the translated geo display
'name.
'@param sh Worksheet. The show/hide worksheet.
'@param dict ILLdictionary. The linelist dictionary.
'@export
Public Sub UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary): End Sub


'@section Admin Code Resolution
'===============================================================================

'@jump:admin-code
'@sub-title Resolve the admin code from a translated display name
'@details
'Compares the supplied translated admin name against GeoNames for each
'admin level and returns the matching raw code ("adm1" through "adm4").
'Falls back to "adm1" when no match is found.
'@param selectedAdmin String. The translated admin display name.
'@return String. Admin code ("adm1", "adm2", "adm3", or "adm4").
'@export
Public Function AdminCode(ByVal selectedAdmin As String) As String: End Function
