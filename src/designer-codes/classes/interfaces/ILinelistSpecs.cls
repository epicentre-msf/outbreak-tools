VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Build")
'@ModuleDescription("Interface of Linelist specification")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods

Public Property Get Dictionary() As ILLdictionary:   End Property
Public Property Get Choices() As ILLChoices:   End Property
Public Property Get specswkb() As Workbook: End Property

'Translation objects (messages, shapes and userforms)
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject: End Property

'Formula data object
Public Property Get FormulaDataObject() As IFormulaData: End Property

Public Property Get TemporarySheetName(ByVal scope As TemporarySheetTypes) As String: End Property
Public Property Get DesignFormat() As ILLFormat: End Property

Public Property Get MainObject() As IMain: End Property
Public Property Get Password() As ILLPasswords: End Property 'Passwords

Public Property Get geoObject() As ILLGeo: End Property      'Geo object
Public Property Get Analysis() As ILLAnalysis: End Property  'Analysis
Public Property Get Value(ByVal tagName As String) As String: End Property 'Get the designer value of some specs

'Exposed methods
'Categories of one variable
Public Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Function

'Void methods
Public Sub Prepare(): End Sub

'Export specifications to workbook
Public Sub Export(ByVal toWkb As Workbook): End Sub
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString): End Sub
