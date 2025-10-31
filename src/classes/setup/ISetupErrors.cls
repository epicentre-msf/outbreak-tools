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

'@section Lifecycle
'===============================================================================
'@sub-title Bind the setup checker to a workbook and optional collaborators.
'@param hostBook Workbook containing setup tables and helper sheets.
'@param passwordService Optional password manager handling sheet protection.
'@param dictionary Optional dictionary wrapper providing data access.
'@param choices Optional choices wrapper exposing list names and categories.
'@param formulaStore Optional formula data handler used to validate expressions.
'@param stateScope Optional application state scope guarding Excel settings.
Public Sub Initialise(ByVal hostBook As Workbook, _
                      Optional ByVal passwordService As IPasswords, _
                      Optional ByVal dictionary As ILLdictionary, _
                      Optional ByVal choices As ILLChoices, _
                      Optional ByVal formulaStore As IFormulaData, _
                      Optional ByVal stateScope As IApplicationState)
End Sub

'@section Execution
'===============================================================================
'@sub-title Execute all setup consistency checks and publish the report.
Public Sub Run()
End Sub

'@section Results
'===============================================================================
'@sub-title Retrieve the collection of generated checking objects.
'@return BetterArray containing the generated IChecking instances.
Public Property Get Checkings() As BetterArray
End Property

