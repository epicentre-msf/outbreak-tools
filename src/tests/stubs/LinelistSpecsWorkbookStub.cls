VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecsWorkbookStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub exposing a workbook-backed ILinelistSpecs implementation for testing")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelistSpecs

Private Type TSpecs
    dictionary As ILLdictionary
    workbook As Workbook
    password As IPasswords
    translations As ITranslationObject
End Type

Private this As TSpecs

Public Sub Initialise(ByVal dictionary As ILLdictionary, _
                      ByVal workbook As Workbook)
    Set this.dictionary = dictionary
    Set this.workbook = workbook
End Sub

Private Sub ILinelistSpecs_ResetCaches()
End Sub

Private Property Get ILinelistSpecs_Workbook() As Workbook
    Set ILinelistSpecs_Workbook = this.workbook
End Property

Private Property Get ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = this.dictionary
End Property

Private Property Get ILinelistSpecs_Choices() As ILLChoices
    Set ILinelistSpecs_Choices = Nothing
End Property

Private Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = Nothing
End Property

Private Property Get ILinelistSpecs_Analysis() As IAnalysis
    Set ILinelistSpecs_Analysis = Nothing
End Property

Private Property Get ILinelistSpecs_ExportObject() As ILLExport
    Set ILinelistSpecs_ExportObject = Nothing
End Property

Private Property Get ILinelistSpecs_Password() As IPasswords
    Set ILinelistSpecs_Password = this.password
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = Nothing
End Property

Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = Nothing
End Property

Private Property Get ILinelistSpecs_HasTemplate() As Boolean
    ILinelistSpecs_HasTemplate = False
End Property

Private Property Get ILinelistSpecs_Translations() As ILLTranslation
    Set ILinelistSpecs_Translations = Nothing
End Property

Private Property Get ILinelistSpecs_DesignerTranslations() As IDesignerTranslation
    Set ILinelistSpecs_DesignerTranslations = Nothing
End Property

Private Property Get ILinelistSpecs_TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
    Set ILinelistSpecs_TransObject = this.translations
End Property

Public Sub SetPassword(ByVal value As IPasswords)
    Set this.password = value
End Sub

Public Sub SetTranslations(ByVal value As ITranslationObject)
    Set this.translations = value
End Sub

Private Property Get ILinelistSpecs_TemporarySheetName(ByVal scope As Byte) As String
    ILinelistSpecs_TemporarySheetName = vbNullString
End Property

Private Property Get ILinelistSpecs_Value(ByVal tagName As String) As String
    ILinelistSpecs_Value = vbNullString
End Property

Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Set ILinelistSpecs_Categories = Nothing
End Function

Private Sub ILinelistSpecs_Prepare(ByVal importService As IDesignerImportService)
End Sub

Private Property Get ILinelistSpecs_OutputWorkbook() As Workbook
    Set ILinelistSpecs_OutputWorkbook = Nothing
End Property

Private Property Get ILinelistSpecs_OutputTemporaryRepos() As ITemporaryRepos
    Set ILinelistSpecs_OutputTemporaryRepos = Nothing
End Property

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

