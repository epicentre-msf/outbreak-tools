VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Build")
'@ModuleDescription("Interface of Linelist specification")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Property Get Dictionary() As ILLdictionary:   End Property
Public Property Get Choices() As ILLchoice:   End Property
Public Property Get Wkb() As Workbook: End Property

Public Property Get TransObject(Optional ByVal loName As String = "T_TradLLMsg") As ITranslation: End Property
'Formula data object
Public Property Get FormDataObject() As IFormulaData: End Property

Public Property Get TemporarySheetName(ByVal scope As TemporarySheetTypes) As String: End Property
Public Property Get DesignFormat() As ILLFormat: End Property

'Exposed methods
'Categories of one variable
Public Function Categories(ByVal varName As String) As BetterArray: End Function

'Void methods
Public Sub Prepare(): End Sub

