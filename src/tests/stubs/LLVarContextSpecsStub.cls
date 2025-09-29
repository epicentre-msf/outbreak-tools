VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarContextSpecsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("Minimal ILinelistSpecs implementation for LLVar tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILinelistSpecs

Private categoriesMap As Collection
Private designFormat As ILLFormat

Private Sub Class_Initialize()
    Set categoriesMap = New Collection
End Sub

Public Sub SetCategories(ByVal varName As String, ByVal values As Variant)
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

    key = LCase$(Trim$(varName))

    On Error Resume Next
        categoriesMap.Remove key
    On Error GoTo 0

    categoriesMap.Add categoryList, key
End Sub

Public Sub SetDesignFormat(ByVal fmt As ILLFormat)
    Set designFormat = fmt
End Sub

Private Function ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = Nothing
End Function

Private Function ILinelistSpecs_Choices() As ILLChoices
    Set ILinelistSpecs_Choices = Nothing
End Function

Private Function ILinelistSpecs_specswkb() As Workbook
    Set ILinelistSpecs_specswkb = Nothing
End Function

Private Function ILinelistSpecs_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set ILinelistSpecs_TransObject = Nothing
End Function

Private Function ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = Nothing
End Function

Private Function ILinelistSpecs_TemporarySheetName(ByVal scope As TemporarySheetTypes) As String
    ILinelistSpecs_TemporarySheetName = vbNullString
End Function

Private Function ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = designFormat
End Function

Private Function ILinelistSpecs_MainObject() As IMain
    Set ILinelistSpecs_MainObject = Nothing
End Function

Private Function ILinelistSpecs_Password() As ILLPasswords
    Set ILinelistSpecs_Password = Nothing
End Function

Private Function ILinelistSpecs_geoObject() As ILLGeo
    Set ILinelistSpecs_geoObject = Nothing
End Function

Private Function ILinelistSpecs_Analysis() As ILLAnalysis
    Set ILinelistSpecs_Analysis = Nothing
End Function

Private Function ILinelistSpecs_Value(ByVal tagName As String) As String
    ILinelistSpecs_Value = vbNullString
End Function

Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim key As String
    Dim stored As Variant
    Dim emptyList As BetterArray

    key = LCase$(Trim$(varName))

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

Private Sub ILinelistSpecs_Prepare()
    'No-op
End Sub

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    'Not required for tests
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    'Not required for tests
End Sub
