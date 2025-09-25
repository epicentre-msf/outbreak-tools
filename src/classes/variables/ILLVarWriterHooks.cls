VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVarWriterHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Extension points for LLVarWriterBase")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'
'@InterfaceDescription("Provides custom behaviour for layout-specific linelist writer needs")

Option Explicit

'@section Layout
'===============================================================================
'@description Configure the supplied layout strategy with workbook anchors.
Public Sub ConfigureLayout(ByVal context As ILLVarContext, _
                           ByVal layout As ILLVarLayout)
End Sub

'@section Behaviour Hooks
'===============================================================================
'@description Invoked after the base writer has populated the label text.
Public Sub AfterLabel(ByVal context As ILLVarContext, _
                      ByVal layout As ILLVarLayout, _
                      ByVal labelCell As Range, _
                      ByVal mainLabelLength As Long)
End Sub

'@description Invoked after the base writer has configured the value type.
Public Sub AfterType(ByVal context As ILLVarContext, _
                     ByVal layout As ILLVarLayout, _
                     ByVal valueCell As Range)
End Sub

'@description Provide any additional formulas unique to the layout.
Public Sub WriteFormulas(ByVal context As ILLVarContext, _
                         ByVal layout As ILLVarLayout)
End Sub

'@description Provide additional choice wiring that is not covered by the base writer.
Public Sub WriteChoices(ByVal context As ILLVarContext, _
                        ByVal layout As ILLVarLayout)
End Sub

'@description Invoked after the base writer configures data validation.
Public Sub AfterValidation(ByVal context As ILLVarContext, _
                           ByVal layout As ILLVarLayout, _
                           ByVal valueCell As Range)
End Sub

'@description Allow callers to run any final customisations at the end of the write.
Public Sub AfterWrite(ByVal context As ILLVarContext, _
                      ByVal layout As ILLVarLayout)
End Sub
