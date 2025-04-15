VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Choice class"


'@Folder("Backend")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule ProcedureNotUsed
'@Interface

Option Explicit

Private Const DEFAULTCHOICESEPARATOR As String = " | "

'Properties
Public Property Get Wksh() As Worksheet: End Property
Public Property Get StartRow() As Long: End Property

Public Property Get StartColumn() As Long: End Property
Public Property Get ChoiceData() As IDataSheet: End Property

Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range: End Property
Public Property Get Categories(ByVal choiceName As String) As BetterArray: End Property
Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean: End Property
Public Property Get ALLChoices() As BetterArray: End Property
Public Property Get PasteTranslatedCategories(ByVal choiceName As String, ByVal trads As ITranslation, _
                                    Optional ByVal separator As String = DEFAULTCHOICESEPARATOR) As String
End Property

Public Property Get PasteCategories(ByVal choiceName As String, _
                                    Optional ByVal separator As String = DEFAULTCHOICESEPARATOR) As String
End Property

'Exposed methods

'Void methods
'Sort the choices sheet
Public Sub Sort(): End Sub

'Add Headers depending on the translation
Public Sub AddHeaders(): End Sub

'Add choice
Public Sub AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray): End Sub

'Add and remove rows from the choice listObject
Public Sub AddRows(): End Sub
Public Sub RemoveRows(): End Sub
