VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Contract for reusable linelist variable writers")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'
'@InterfaceDescription("Defines the lifecycle of reusable linelist variable writers")

Option Explicit

'@section Lifecycle
'===============================================================================
'@description Prepare the writer with its shared context, layout strategy, and behaviour hooks.
Public Sub Initialise(ByVal context As ILLVarContext, _
                      ByVal layout As ILLVarLayout, _
                      ByVal hooks As ILLVarWriterHooks)
End Sub

'@description Reset the writer to an uninitialised state.
Public Sub Reset()
End Sub

'@section Operations
'===============================================================================
'@description Materialise the variable using the configured context and layout.
Public Sub WriteVariable()
End Sub

