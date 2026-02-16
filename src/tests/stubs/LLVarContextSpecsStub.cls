VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarContextSpecsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal ILinelistSpecs implementation for sections and variable writer tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILinelistSpecs

Private categoriesMap As Collection
Private designFormat As ILLFormat
Private dictionary As ILLdictionary
Private formulaData As IFormulaData
Private translationObject As ITranslationObject
Private geoObject As ILLGeo

Private Sub Class_Initialize()
    Set categoriesMap = New Collection
End Sub

Public Sub SetCategories(ByVal varName As String, values As Variant, _
                         Optional ByVal useShortlabels As Boolean = False)
    Dim categoryList As BetterArray
    Dim idx As Long
    Dim key As String

    If IsArray(values) Then
        Set categoryList = New BetterArray
        categoryList.LowerBound = 1
        For idx = LBound(values) To UBound(values)
            categoryList.Push values(idx)
        Next idx
    Else
        Set categoryList = New BetterArray
        categoryList.LowerBound = 1
        categoryList.Push values
    End If

    key = BuildCategoryKey(varName, useShortlabels)

    On Error Resume Next
        categoriesMap.Remove key
    On Error GoTo 0

    categoriesMap.Add categoryList, key
End Sub

Public Sub SetDesignFormat(ByVal fmt As ILLFormat)
    Set designFormat = fmt
End Sub

Public Sub SetDictionary(ByVal dict As ILLdictionary)
    Set dictionary = dict
End Sub

Public Sub SetFormulaData(ByVal data As IFormulaData)
    Set formulaData = data
End Sub

Public Sub SetTranslation(ByVal translation As ITranslationObject)
    Set translationObject = translation
End Sub

Public Sub SetGeo(ByVal geo As ILLGeo)
    Set geoObject = geo
End Sub

'@section ILinelistSpecs implementation
'===============================================================================

Private Sub ILinelistSpecs_ResetCaches()
    'No-op
End Sub

Private Property Get ILinelistSpecs_Workbook() As Workbook
    Set ILinelistSpecs_Workbook = Nothing
End Property

Private Property Get ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = dictionary
End Property

Private Property Get ILinelistSpecs_Choices() As ILLChoices
    Set ILinelistSpecs_Choices = Nothing
End Property

Private Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = designFormat
End Property

Private Property Get ILinelistSpecs_Analysis() As IAnalysis
    Set ILinelistSpecs_Analysis = Nothing
End Property

Private Property Get ILinelistSpecs_ExportObject() As ILLExport
    Set ILinelistSpecs_ExportObject = Nothing
End Property

Private Property Get ILinelistSpecs_Password() As IPasswords
    Set ILinelistSpecs_Password = Nothing
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = geoObject
End Property

Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = formulaData
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
    Set ILinelistSpecs_TransObject = translationObject
End Property

Private Property Get ILinelistSpecs_TemporarySheetName(ByVal scope As Byte) As String
    ILinelistSpecs_TemporarySheetName = vbNullString
End Property

Private Property Get ILinelistSpecs_Value(ByVal tagName As String) As String
    ILinelistSpecs_Value = vbNullString
End Property

Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim key As String
    Dim stored As Variant
    Dim emptyList As BetterArray

    key = BuildCategoryKey(varName, useShortlabels)

    On Error Resume Next
        Set stored = categoriesMap.Item(key)
    On Error GoTo 0

    If TypeName(stored) = "BetterArray" Then
        Set ILinelistSpecs_Categories = stored
    Else
        Set emptyList = New BetterArray
        emptyList.LowerBound = 1
        Set ILinelistSpecs_Categories = emptyList
    End If
End Function

Private Function BuildCategoryKey(ByVal varName As String, ByVal useShortlabels As Boolean) As String
    BuildCategoryKey = LCase$(Trim$(varName)) & "|" & CStr(useShortlabels)
End Function

Private Sub ILinelistSpecs_Prepare(ByVal importService As IDesignerImportService)
    'No-op
End Sub

Private Property Get ILinelistSpecs_OutputWorkbook() As Workbook
    Set ILinelistSpecs_OutputWorkbook = Nothing
End Property

Private Property Get ILinelistSpecs_OutputTemporaryRepos() As ITemporaryRepos
    Set ILinelistSpecs_OutputTemporaryRepos = Nothing
End Property

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    'Not required for tests
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    'Not required for tests
End Sub

