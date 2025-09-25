VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRFVarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Factory building CRF variable writers")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Function Create(ByVal context As ILLVarContext) As ILLVarWriter
    Dim layout As ILLVarLayout
    Dim hooks As ILLVarWriterHooks
    Dim writer As ILLVarWriter

    If context Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, "CRFVarWriter", _
                  "Context instance is required when creating a CRF writer"
    End If

    Set layout = New LLVarLayoutCRF
    Set hooks = New LLVarWriterHooksCRF
    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks

    Set Create = writer
End Function

