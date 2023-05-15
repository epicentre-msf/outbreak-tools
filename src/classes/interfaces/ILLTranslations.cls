VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of Linelist Translation")
'@IgnoreModule

Option Explicit

Public Enum TradScopes
    TranslationOfDictionary = 1
    TranslationOfShapes = 2
    TranslationOfForms = 3
    TranslationOfMessages = 4
    TranslationOfRanges = 5
    TranslationOfRibbon = 6
End Enum

'Exposed methods
Public Property Get TransObject(Optional ByVal trans As TradScopes = TranslationOfMessages) As ITranslation: End Property
Public Sub ExportToWkb(ByVal outputWkb As Workbook): End Sub

Public Sub ImportTranslation(ByVal fromWkb As Workbook): End Sub
