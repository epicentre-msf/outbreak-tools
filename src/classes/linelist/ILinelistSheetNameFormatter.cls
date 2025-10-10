VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSheetNameFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Centralises sheet-name construction logic for the Linelist workflow"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Centralises sheet-name construction logic for the Linelist workflow")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Scope enumerations
'===============================================================================
Public Enum LinelistSheetScope
    sheetScopeStandard = 1
    sheetScopePrint = 2
    sheetScopeCrf = 3
End Enum

Public Enum LinelistCodeScope
    codeScopeClass = 1
    codeScopeModule = 2
End Enum


'@section Configuration
'===============================================================================
'@sub-title Prefix assigned to printed worksheets.
Public Property Get PrintPrefix() As String
End Property

'@sub-title Prefix assigned to CRF worksheets.
Public Property Get CrfPrefix() As String
End Property

'@sub-title Maximum supported sheet-name length including prefixes.
Public Property Get SheetNameMaxLength() As Long
End Property


'@section Sheet helpers
'===============================================================================
'@sub-title Resolve the scoped sheet name within the allowed length budget.
'@param baseName Candidate workbook sheet name.
'@param scope Sheet naming scope (standard, print, crf).
'@return Sheet name trimmed to the configured maximum with the scope prefix applied.
Public Function FormatSheetName(ByVal baseName As String, _
                                Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As String
End Function

'@sub-title Check whether a scoped sheet exists in the workbook.
Public Function SheetExists(ByVal workbook As Workbook, ByVal baseName As String, _
                            Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
End Function

'@sub-title Retrieve a scoped worksheet or raise if not found.
Public Function ResolveWorksheet(ByVal workbook As Workbook, ByVal baseName As String, _
                                 Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
End Function
