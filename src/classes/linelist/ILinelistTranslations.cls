VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the improved linelist translations manager"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface describing the improved linelist translations manager")
'@depends ITranslationObject
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName

Option Explicit

'@section Enumerations
'===============================================================================
'@details Strongly typed scopes describing supported translation collections.
Public Enum LinelistTranslationScope
    translationScopeDictionary = 1
    translationScopeShapes = 2
    translationScopeForms = 3
    translationScopeMessages = 4
    translationScopeRibbon = 5
End Enum

'@section Core contract
'===============================================================================
'@sub-title Retrieve a translation object for the requested scope.
'@param scope LinelistTranslationScope describing the table to load.
'@return ITranslationObject wrapping the underlying listobject.
Public Property Get TransObject(Optional ByVal scope As Byte = translationScopeMessages) As ITranslationObject
End Property

'@sub-title Retrieve a named value exposed on the translation worksheet.
'@param tagName Lowercase tag identifying the named range to read.
'@return String value stored inside the named range.
Public Property Get Value(ByVal tagName As String) As String
End Property

'@sub-title Update a named value exposed on the translation worksheet.
'@param tagName Lowercase tag identifying the named range to write.
'@param newValue String value to assign to the named range.
Public Sub SetValue(ByVal tagName As String, ByVal newValue As String)
End Sub

'@sub-title Export translations to the target workbook honouring supplied options.
'@param target Workbook where dedicated translation sheets will be created.
'@param visibility Optional XlSheetVisibility controlling post-export visibility.
Public Sub ExportToWorkbook(ByVal target As Workbook, Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)
End Sub

'@sub-title Import translations from the source workbook.
'@param source Workbook containing dictionary translations to synchronise.
Public Sub ImportFromWorkbook(ByVal source As Workbook)
End Sub

'@sub-title Export setup dictionary translations to a workbook for round-tripping.
Public Sub ExportSetup(ByVal target As Workbook, Optional ByVal visibility As XlSheetVisibility = xlSheetVeryHidden)
End Sub

'@sub-title Import setup dictionary translations from a worksheet containing a single table.
'@param sourceSheet Worksheet whose first ListObject represents the dictionary translation.
Public Sub ImportFromSetup(ByVal sourceSheet As Worksheet)
End Sub
