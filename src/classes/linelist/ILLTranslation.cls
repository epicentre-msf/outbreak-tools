VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for managing linelist translation tables"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract for managing linelist translation tables")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class ILLTranslation
'@description
'Defines the public contract for the LLTranslation class, exposing
'scoped translation objects, import/export of translation ListObjects,
'dictionary exchange, and workbook-level hidden name initialisation
'for translated configuration values.


'@section Enumerations
'===============================================================================
'@description Supported translation content scopes.

'@sub-title Available translation content scopes
'@details
'Each value maps to a specific ListObject on the translation worksheet.
'Used as a Byte parameter to select which translation table backs the
'returned ITranslationObject.
Public Enum TranslationScope
    TranslationOfMessages = 0
    TranslationOfDictionary = 1
    TranslationOfShapes = 2
    TranslationOfForms = 3
    TranslationOfRibbon = 4
    TranslationOfRanges = 5
    TranslationOfDropdowns = 6
End Enum


'@section Properties
'===============================================================================
'@description Accessors for translation objects and the host worksheet.

'@jump:trans-object
'@prop-title Scoped translation object
'@details
'Returns an ITranslationObject configured for the requested translation
'scope. The language code is resolved from the host workbook HiddenNames.
'@param trans Optional Byte. TranslationScope value selecting the content
'scope. Defaults to TranslationOfMessages.
'@return ITranslationObject. The translation object for the requested scope.
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
End Property

'@jump:wksh
'@prop-title Host translation worksheet
'@details
'Returns the worksheet bound to this LLTranslation instance, hosting
'all translation ListObjects.
'@return Worksheet. The host translation worksheet.
Public Property Get Wksh() As Worksheet
End Property


'@section DataExchange
'===============================================================================
'@description Import, export, and dictionary exchange operations.

'@jump:export
'@sub-title Export all translation tables to an output workbook
'@details
'Creates or clears a worksheet in the output workbook matching the host
'sheet name, then exports every ListObject from the host sheet using
'CustomTable.Export with stacking.
'@param outputWkb Workbook. The destination workbook.
'@param Hide Optional Long. Sheet visibility after export. Defaults to xlSheetVeryHidden.
Public Sub Export(ByVal outputWkb As Workbook, _
                  Optional ByVal Hide As Long = xlSheetVeryHidden)
End Sub

'@jump:import
'@sub-title Import translation tables from another workbook
'@details
'Finds a worksheet in the source workbook matching the host sheet name.
'Cleans the dictionary ListObject before import. For each ListObject on
'the host sheet, imports data from the matching ListObject on the source
'sheet. Missing source tables are silently skipped.
'@param fromWkb Workbook. The source workbook containing translation data.
Public Sub Import(ByVal fromWkb As Workbook)
End Sub

'@jump:import-dictionary
'@sub-title Import dictionary translations from a raw ListObject
'@details
'Overwrites the host dictionary table with data from the supplied
'ListObject, preserving the source column headers.
'@param dictionaryTable ListObject. The source dictionary ListObject.
Public Sub ImportDictionary(ByVal dictionaryTable As ListObject)
End Sub

'@jump:export-dictionary
'@sub-title Export the dictionary table to a named worksheet
'@details
'Creates or clears a worksheet in the output workbook with the given
'name, then exports the dictionary ListObject using CustomTable.Export.
'@param outputWkb Workbook. The destination workbook.
'@param sheetName Optional String. Worksheet name for the export. Defaults to "Translations".
'@param Hide Optional Long. Sheet visibility after export. Defaults to xlSheetVisible.
Public Sub ExportDictionary(ByVal outputWkb As Workbook, _
                            Optional ByVal sheetName As String = "Translations", _
                            Optional ByVal Hide As Long = xlSheetVisible)
End Sub


'@section HiddenNames
'===============================================================================
'@description Workbook-level hidden name initialisation from translations.

'@jump:initialise-hidden-names
'@sub-title Create RNG_* hidden names on a target workbook
'@details
'Translates message and sheet-name tags using the TranslationOfMessages
'scope, then writes each translated value as a workbook-level hidden name
'on the target workbook. These names replace the legacy named ranges that
'were previously stored on the translation worksheet.
'@param targetWkb Workbook. The workbook receiving the hidden names.
Public Sub InitialiseHiddenNames(ByVal targetWkb As Workbook)
End Sub
