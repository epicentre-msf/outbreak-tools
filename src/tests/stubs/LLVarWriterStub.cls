VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("Minimal ILLVarWriter implementation for tests")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriter

Private context As ILLVarContext
Private layout As ILLVarLayout
Private hooks As ILLVarWriterHooks

Public Writes As Long

Private Sub ILLVarWriter_Initialise(ByVal ctx As ILLVarContext, _
                                    ByVal lay As ILLVarLayout, _
                                    ByVal hk As ILLVarWriterHooks)
    Set context = ctx
    Set layout = lay
    Set hooks = hk
End Sub

Private Sub ILLVarWriter_Reset()
    Set context = Nothing
    Set layout = Nothing
    Set hooks = Nothing
End Sub

Private Sub ILLVarWriter_WriteVariable()
    Writes = Writes + 1
End Sub

