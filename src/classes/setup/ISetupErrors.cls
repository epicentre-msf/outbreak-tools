VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the setup consistency checker"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Interface describing the setup consistency checker")
'@depends ProjectError, BetterArray, IChecking, ICheckingOutput
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class ISetupErrors
'@description
'Defines the public contract for the SetupErrors class, exposing methods
'to initialise a setup consistency checker, execute all validation checks
'(dictionary, choices, exports, translations), and retrieve the resulting
'collection of IChecking report objects.


'@section Lifecycle
'===============================================================================

'@jump:initialise
'@sub-title Bind the setup checker to a workbook and optional collaborators
'@details
'Stores the workbook reference and optional collaborators (password manager,
'dictionary, choices, formula store, application state). When optional
'parameters are omitted, the checker creates default instances internally
'during ConfigureDependencies.
'@param hostBook Workbook. The workbook containing setup tables and helper sheets.
'@param passwordService Optional IPasswords. Password manager handling sheet protection.
'@param dictionary Optional ILLdictionary. Dictionary wrapper providing data access.
'@param choices Optional ILLChoices. Choices wrapper exposing list names and categories.
'@param formulaStore Optional IFormulaData. Formula data handler used to validate expressions.
'@param stateScope Optional IApplicationState. Application state scope guarding Excel settings.
'@export
Public Sub Initialise(ByVal hostBook As Workbook, _
                      Optional ByVal passwordService As IPasswords, _
                      Optional ByVal dictionary As ILLdictionary, _
                      Optional ByVal choices As ILLChoices, _
                      Optional ByVal formulaStore As IFormulaData, _
                      Optional ByVal stateScope As IApplicationState)
End Sub

'@section Execution
'===============================================================================

'@jump:run
'@sub-title Execute all setup consistency checks and publish the report
'@details
'Suspends Excel UI, resets the internal check container, then runs each
'checker in sequence (dictionary, choices, exports, translations). On
'completion, publishes all collected IChecking objects to the reporting
'worksheet. Restores Excel state even when an error occurs.
'@export
Public Sub Run()
End Sub

'@section Results
'===============================================================================

'@jump:checkings
'@prop-title Collection of generated checking objects
'@details
'Returns a BetterArray containing the IChecking instances produced by the
'most recent Run call. Each entry covers one validation domain (dictionary,
'choices, exports, translations).
'@return BetterArray. The generated IChecking instances.
'@export
Public Property Get Checkings() As BetterArray
End Property
