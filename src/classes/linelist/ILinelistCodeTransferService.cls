VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistCodeTransferService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Facade coordinating Linelist code module and form transfers"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Facade coordinating Linelist code module and form transfers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

Public Sub TransferComponent(ByVal componentName As String, _
                             ByVal scope As LinelistCodeScope, _
                             ByVal source As Workbook, _
                             ByVal target As Workbook)
End Sub

Public Sub TransferWorkbookModule(ByVal moduleName As String, _
                                  ByVal source As Workbook, _
                                  ByVal target As Workbook)
End Sub

Public Sub TransferForm(ByVal formName As String, _
                        ByVal source As Workbook, _
                        ByVal target As Workbook)
End Sub

Public Sub TransferComponents(ByVal componentNames As BetterArray, _
                              ByVal scope As LinelistCodeScope, _
                              ByVal source As Workbook, _
                              ByVal target As Workbook)
End Sub

Public Property Get StrategyName() As String
End Property
