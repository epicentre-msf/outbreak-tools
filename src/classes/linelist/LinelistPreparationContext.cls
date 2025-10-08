VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistPreparationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Holds shared state while the Linelist preparation pipeline executes"

'@Folder("Linelist")
'@ModuleDescription("Holds shared state while the Linelist preparation pipeline executes")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@depends ProjectError

Option Explicit

Implements ILinelistPreparationContext

'@section State and constants
'===============================================================================
Private Type TContext
    linelist As ILinelist
    specs As ILinelistSpecs
    dictionary As ILLdictionary
    sheetFormatter As ILinelistSheetNameFormatter
    stateScope As ILinelistApplicationStateScope
End Type

Private this As TContext
Private Const CLASSNAME As String = "LinelistPreparationContext"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal linelist As ILinelist, _
                       ByVal specs As ILinelistSpecs, _
                       ByVal dictionary As ILLdictionary, _
                       ByVal formatter As ILinelistSheetNameFormatter, _
                       ByVal scope As ILinelistApplicationStateScope) As ILinelistPreparationContext

    Dim context As LinelistPreparationContext
    Set context = New LinelistPreparationContext
    context.Initialise linelist, specs, dictionary, formatter, scope
    Set Create = context
End Function

Friend Sub Initialise(ByVal linelist As ILinelist, _
                      ByVal specs As ILinelistSpecs, _
                      ByVal dictionary As ILLdictionary, _
                      ByVal formatter As ILinelistSheetNameFormatter, _
                      ByVal scope As ILinelistApplicationStateScope)

    ValidateLinelist linelist
    ValidateSpecs specs
    ValidateDictionary dictionary
    ValidateFormatter formatter
    ValidateScope scope

    Set this.linelist = linelist
    Set this.specs = specs
    Set this.dictionary = dictionary
    Set this.sheetFormatter = formatter
    Set this.stateScope = scope
End Sub


'@section Validation helpers
'===============================================================================
Private Sub ValidateLinelist(ByVal linelist As ILinelist)
    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist reference cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpecs(ByVal specs As ILinelistSpecs)
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications reference cannot be Nothing"
    End If
End Sub

Private Sub ValidateDictionary(ByVal dictionary As ILLdictionary)
    If dictionary Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dictionary reference cannot be Nothing"
    End If
End Sub

Private Sub ValidateFormatter(ByVal formatter As ILinelistSheetNameFormatter)
    If formatter Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Sheet formatter reference cannot be Nothing"
    End If
End Sub

Private Sub ValidateScope(ByVal scope As ILinelistApplicationStateScope)
    If scope Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application state scope cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Property Get ILinelistPreparationContext_Linelist() As ILinelist
    Set ILinelistPreparationContext_Linelist = this.linelist
End Property

Private Property Get ILinelistPreparationContext_Specifications() As ILinelistSpecs
    Set ILinelistPreparationContext_Specifications = this.specs
End Property

Private Property Get ILinelistPreparationContext_Dictionary() As ILLdictionary
    Set ILinelistPreparationContext_Dictionary = this.dictionary
End Property

Private Property Get ILinelistPreparationContext_SheetFormatter() As ILinelistSheetNameFormatter
    Set ILinelistPreparationContext_SheetFormatter = this.sheetFormatter
End Property

Private Property Get ILinelistPreparationContext_StateScope() As ILinelistApplicationStateScope
    Set ILinelistPreparationContext_StateScope = this.stateScope
End Property

