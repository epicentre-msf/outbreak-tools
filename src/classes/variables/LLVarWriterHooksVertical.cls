VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterHooksVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Default hooks for vertical variable writer")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriterHooks

Private Sub ILLVarWriterHooks_AfterLabel(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout, _
                                         ByVal labelCell As Range, _
                                         ByVal mainLabelLength As Long)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_AfterType(ByVal context As ILLVarContext, _
                                        ByVal layout As ILLVarLayout, _
                                        ByVal valueCell As Range)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_AfterValidation(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout, _
                                              ByVal valueCell As Range)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_AfterWrite(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_ConfigureLayout(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout)
    Dim anchorRow As Long
    Dim sheet As Worksheet

    anchorRow = CLng(context.ValueOf("column index"))
    Set sheet = context.Worksheet

    layout.Configure sheet, anchorRow
End Sub

Private Sub ILLVarWriterHooks_WriteChoices(ByVal context As ILLVarContext, _
                                           ByVal layout As ILLVarLayout)
    'Default writer already handles basic dropdown wiring
End Sub

Private Sub ILLVarWriterHooks_WriteFormulas(ByVal context As ILLVarContext, _
                                            ByVal layout As ILLVarLayout)
    'Not required for default vertical variables
End Sub

