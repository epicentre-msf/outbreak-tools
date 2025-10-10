VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVarContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Shared context information for linelist variable writers")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'
'@InterfaceDescription("Provides cached access to linelist metadata used when materialising variables")

Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Initialise(ByVal ll As ILinelist, _
                      ByVal varName As String, _
                      Optional ByVal dictionary As ILLdictionary, _
                      Optional ByVal variables As ILLVariables, _
                      Optional ByVal specs As ILinelistSpecs)
End Sub

Public Property Get Linelist() As ILinelist
End Property

Public Property Get VariableName() As String
End Property

Public Sub ResetCaches()
End Sub

'@section Metadata lookups
'===============================================================================
Public Function ValueOf(ByVal columnName As String) As String
End Function

Public Function Dictionary() As ILLdictionary
End Function

Public Function Variables() As ILLVariables
End Function

Public Function Specs() As ILinelistSpecs
End Function

Public Function Design() As ILLFormat
End Function

Public Function Translation() As ITranslationObject
End Function

Public Function Categories(ByVal varName As String, _
                           Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Function

Public Function EditableLabelFlag() As String
End Function

Public Function RegisterBookDirective() As String
End Function

Public Function FormattingConditionName() As String
End Function

'@section Collaborators
'===============================================================================
Public Function Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
End Function

Public Function Worksheet() As Worksheet
End Function

Public Function PrintWorksheet() As Worksheet
End Function

