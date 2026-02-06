VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Contract describing reusable linelist specifications accessors")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@InterfaceDescription("Surface area consumed by linelist writers and contexts to retrieve workbook driven specifications")

Option Explicit

'@section Lifecycle
'===============================================================================
'@description Refresh any cached collaborators so subsequent reads query the workbook again.
Public Sub ResetCaches()
End Sub

'@description Workbook hosting the specifications.
Public Property Get Workbook() As Workbook
End Property

'@section Core Specifications
'===============================================================================
Public Property Get Dictionary() As ILLdictionary
End Property

Public Property Get Choices() As ILLChoices
End Property

Public Property Get DesignFormat() As ILLFormat
End Property

Public Property Get Analysis() As ILLAnalysis
End Property

Public Property Get ExportObject() As ILLExport
End Property

Public Property Get Password() As IPasswords
End Property

Public Property Get GeoObject() As ILLGeo
End Property

Public Property Get FormulaDataObject() As IFormulaData
End Property

Public Property Get MainObject() As IMain
End Property

Public Property Get Translations() As ILLTranslations
End Property

Public Property Get DesignerTranslations() As IDesTranslation
End Property

Public Property Get TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
End Property

'@section Metadata Helpers
'===============================================================================
Public Property Get TemporarySheetName(ByVal scope As Byte) As String
End Property

Public Property Get Value(ByVal tagName As String) As String
End Property

Public Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Function

'@section Workbook Management
'===============================================================================
Public Sub Prepare()
End Sub

Public Sub Export(ByVal toWkb As Workbook)
End Sub

Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

