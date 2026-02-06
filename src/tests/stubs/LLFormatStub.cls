VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLFormatStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("No-op ILLFormat implementation for unit tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLFormat

Private Sub ILLFormat_ApplyFormat(formattingObject As Variant, ByVal scope As Byte, _
                                  Optional ByVal nbDecimals As Integer = 0)
    'No-op: silently ignore all format requests during tests
End Sub

Private Function ILLFormat_DesignValue(ByVal labelName As String, _
                                       Optional ByVal returnedColor As Boolean = True) As Variant
    ILLFormat_DesignValue = 0
End Function

Private Sub ILLFormat_Import(ByVal inpsh As Worksheet)
    'Not required for tests
End Sub

Private Sub ILLFormat_Export(ByVal toWkb As Workbook)
    'Not required for tests
End Sub

Private Property Get ILLFormat_HasCheckings() As Boolean
    ILLFormat_HasCheckings = False
End Property

Private Property Get ILLFormat_CheckingValues() As IChecking
    Set ILLFormat_CheckingValues = Nothing
End Property
