VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMasterSetupVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Master Setup Variables manager"

'@Interface
'@Folder("MasterSetup")
'@ModuleDescription("Interface for Master Setup Variables manager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal dropdowns As IDropdownLists)
End Sub

Public Sub RefreshChoices(ByVal variableName As String, ByVal choices As ILLChoices)
End Sub

'@section Lookups
'===============================================================================
Public Function LabelFor(ByVal variableName As String) As String
End Function

Public Function DefaultChoiceFor(ByVal variableName As String) As String
End Function

Public Function ChoicesFor(ByVal variableName As String) As String
End Function

Public Function DefaultStatusFor(ByVal variableName As String) As String
End Function

Public Function SectionFor(ByVal variableName As String) As String
End Function

'@section Row management
'===============================================================================
Public Sub AddRows(Optional ByVal count As Long = 5)
End Sub

Public Sub RemoveRows(Optional ByVal count As Long = 1)
End Sub

'@section Translations
'===============================================================================
Public Sub TranslateHeaders(ByVal trads As ITranslationObject)
End Sub

Public Sub TranslateContent(ByVal trads As ITranslationObject)
End Sub

Public Sub TranslateAll(ByVal trads As ITranslationObject)
End Sub

'@section Export / Import
'===============================================================================
Public Function CloneToWorkbook(ByVal targetWorkbook As Workbook, _
                                Optional ByVal sheetName As String = vbNullString, _
                                Optional ByVal startCell As Range = Nothing) As IMasterSetupVariables
End Function

Public Sub RestoreFrom(ByVal source As IMasterSetupVariables)
End Sub

Public Sub ExportToWorksheet(ByVal targetSheet As Worksheet)
End Sub

Public Sub ImportFromWorksheet(ByVal sourceSheet As Worksheet)
End Sub

Public Sub CopyToListObject(ByVal targetTable As ListObject)
End Sub

Public Sub ExportMetadata(ByVal targetSheet As Worksheet)
End Sub

'@section Accessors
'===============================================================================
Public Property Get Table() As ListObject
End Property

Public Property Get Initialised() As Boolean
End Property

'@sub-title Return the underlying range for a specific column header.
Public Function DataRange(ByVal headerName As String, _
                          Optional ByVal includeHeaders As Boolean = False) As Range
End Function

'@section Validation
'===============================================================================
'@sub-title Apply workbook-scoped variable dropdown validation to a target range.
Public Sub SetVariableValidation(ByVal targetRange As Range)
End Sub

