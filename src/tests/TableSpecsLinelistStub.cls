VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsLinelistStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal ILinelistSpecs double that exposes a configurable dictionary for TableSpecs tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ILinelistSpecs

Private Type TLinelistStub
    dictionary As ILLdictionary
End Type

Private this As TLinelistStub

'@section Configuration
'===============================================================================
'Configure the dictionary returned through the interface.
Public Sub SetDictionary(ByVal dictionary As ILLdictionary)
    Set this.dictionary = dictionary
End Sub

'@section ILinelistSpecs implementation
'===============================================================================
Private Sub ILinelistSpecs_ResetCaches()
End Sub

Private Property Get ILinelistSpecs_Workbook() As Workbook
    Set ILinelistSpecs_Workbook = Nothing
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

Private Property Get ILinelistSpecs_Analysis() As ILLAnalysis
    Set ILinelistSpecs_Analysis = Nothing
End Property

Private Property Get ILinelistSpecs_ExportObject() As ILLExport
    Set ILinelistSpecs_ExportObject = Nothing
End Property

Private Property Get ILinelistSpecs_Password() As ILLPasswords
    Set ILinelistSpecs_Password = Nothing
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = Nothing
End Property

Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = Nothing
End Property

Private Property Get ILinelistSpecs_MainObject() As IMain
    Set ILinelistSpecs_MainObject = Nothing
End Property

Private Property Get ILinelistSpecs_Translations() As ILLTranslations
    Set ILinelistSpecs_Translations = Nothing
End Property

Private Property Get ILinelistSpecs_DesignerTranslations() As IDesTranslation
    Set ILinelistSpecs_DesignerTranslations = Nothing
End Property

Private Property Get ILinelistSpecs_TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
    Set ILinelistSpecs_TransObject = Nothing
End Property

Private Property Get ILinelistSpecs_TemporarySheetName(ByVal scope As Byte) As String
    ILinelistSpecs_TemporarySheetName = vbNullString
End Property

Private Property Get ILinelistSpecs_Value(ByVal tagName As String) As String
    ILinelistSpecs_Value = vbNullString
End Property

Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim emptyList As BetterArray
    Set emptyList = New BetterArray
    emptyList.LowerBound = 1
    Set ILinelistSpecs_Categories = emptyList
End Function

Private Sub ILinelistSpecs_Prepare()
End Sub

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

Private Function ILinelistSpecs_CreateSectionContext(ByVal linelist As ILinelist, ByVal startRow As Long) As ILLSectionContext
    Set ILinelistSpecs_CreateSectionContext = Nothing
End Function

Private Function ILinelistSpecs_CreateSectionWriter() As ILLSectionWriter
    Set ILinelistSpecs_CreateSectionWriter = Nothing
End Function

Private Function ILinelistSpecs_CreateSubSectionWriter() As ILLSubSectionWriter
    Set ILinelistSpecs_CreateSubSectionWriter = Nothing
End Function

Private Function ILinelistSpecs_CreateVarContext(ByVal linelist As ILinelist, ByVal variableName As String) As ILLVarContext
    Set ILinelistSpecs_CreateVarContext = Nothing
End Function

Private Function ILinelistSpecs_CreateVarWriter() As ILLVarWriter
    Set ILinelistSpecs_CreateVarWriter = Nothing
End Function
