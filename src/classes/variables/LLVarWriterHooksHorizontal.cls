VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterHooksHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Default hooks for horizontal variable writer")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriterHooks

Private Sub ILLVarWriterHooks_AfterLabel(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout, _
                                         ByVal labelCell As Range, _
                                         ByVal mainLabelLength As Long)
    'Formatting handled elsewhere
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
    Dim anchorColumn As Long
    Dim sheet As Worksheet
    Dim printSheet As Worksheet

    anchorColumn = CLng(context.ValueOf("column index"))
    Set sheet = context.Worksheet
    Set printSheet = Nothing

    On Error Resume Next
        Set printSheet = context.PrintWorksheet
    On Error GoTo 0

    layout.Configure sheet, anchorColumn, printSheet
End Sub

Private Sub ILLVarWriterHooks_WriteChoices(ByVal context As ILLVarContext, _
                                           ByVal layout As ILLVarLayout)
    'Default writer already handles basic dropdown wiring
End Sub

Private Sub ILLVarWriterHooks_WriteFormulas(ByVal context As ILLVarContext, _
                                            ByVal layout As ILLVarLayout)
    'Not required for default horizontal variables
End Sub

