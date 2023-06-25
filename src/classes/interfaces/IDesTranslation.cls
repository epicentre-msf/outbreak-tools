VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Designer")
'@ModuleDescription("Interface of translation Object for the designer")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslation: End Property
Public Sub TranslateDesigner(ByVal mainsh As Worksheet): End Sub
Public Property Get LangListRng() As Range: End Property
'Get translated value of a message code
Public Function TranslationMsg(ByVal msgCode As String) As String: End Function

