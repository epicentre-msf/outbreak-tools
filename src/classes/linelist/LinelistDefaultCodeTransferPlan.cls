VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistDefaultCodeTransferPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides the canonical list of VBA components required by the Linelist workbook"

'@Folder("Linelist")
'@ModuleDescription("Provides the canonical list of VBA components required by the Linelist workbook")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistCodeTransferPlan

'@section State and constants
'===============================================================================
Private Type TPlan
    classModules As BetterArray
    standardModules As BetterArray
    formNames As BetterArray
    workbookModuleName As String
End Type

Private this As TPlan
Private Const CLASSNAME As String = "LinelistDefaultCodeTransferPlan"
Private Const DEFAULT_WORKBOOK_MODULE As String = "LinelistWorkbook"


'@section Factory helpers
'===============================================================================
Public Function Create() As ILinelistCodeTransferPlan
    Dim plan As LinelistDefaultCodeTransferPlan
    Set plan = New LinelistDefaultCodeTransferPlan
    plan.Initialise
    Set Create = plan
End Function

Friend Sub Initialise()
    Set this.classModules = BuildArray(DefaultClassModules())
    Set this.standardModules = BuildArray(DefaultStandardModules())
    Set this.formNames = BuildArray(DefaultFormNames())
    this.workbookModuleName = DEFAULT_WORKBOOK_MODULE
End Sub


'@section Defaults
'===============================================================================
Private Function DefaultClassModules() As Variant
    DefaultClassModules = Array( _
        "BetterArray", "DataSheet", "IDataSheet", "LLExport", _
        "ILLExport", "LLChoices", "ILLChoices", _
        "LLdictionary", "ILLdictionary", "LLGeo", _
        "ILLGeo", "Passwords", "IPasswords", _
        "LLSheets", "ILLSheets", "LLTranslations", _
        "ILLTranslations", "OSFiles", "IOSFiles", _
        "TranslationObject", "ITranslationObject", "LLVariables", _
        "ILLVariables", "ILLSpatial", "LLSpatial", _
        "DropdownLists", "IDropdownLists", "LLShowHide", _
        "ILLShowHide", "CustomTable", "ICustomTable", _
        "OutputSpecs", "IOutputSpecs", _
        "ImpRep", "IImpRep", "ImpSpecs", "IImpSpecs", _
        "UpVal", "IUpVal", "IGraphs", "Graphs", "AnaTabIds", _
        "IAnaTabIds", "ExportButtons", "IExportButtons" _
    )
End Function

Private Function DefaultStandardModules() As Variant
    DefaultStandardModules = Array( _
        "LinelistCustomFunctions", _
        "LinelistGeo", "LinelistEvents", _
        "EventsSpatialAnalysis", _
        "EventsLinelistButtons", "EventsGlobalAnalysis" _
    )
End Function

Private Function DefaultFormNames() As Variant
    DefaultFormNames = Array( _
        "F_Advanced", "F_Export", "F_ExportMig", _
        "F_Geo", "F_ImportRep", _
        "F_ShowHideLL", "F_ShowHidePrint", "F_ShowVarLabels", _
        "F_EpiWeek" _
    )
End Function


'@section Helpers
'===============================================================================
Private Function BuildArray(ByVal values As Variant) As BetterArray
    Dim result As BetterArray
    Dim index As Long

    Set result = New BetterArray
    result.LowerBound = 1

    For index = LBound(values) To UBound(values)
        result.Push CStr(values(index))
    Next index

    Set BuildArray = result
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub

Private Sub EnsureArray(ByVal candidate As BetterArray, ByVal description As String)
    If candidate Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, description & " cannot be Nothing"
    End If
End Sub


'@section Interface implementation
'===============================================================================
Private Function ILinelistCodeTransferPlan_ClassComponentNames() As BetterArray
    EnsureArray this.classModules, "Class component list"
    Set ILinelistCodeTransferPlan_ClassComponentNames = this.classModules
End Function

Private Function ILinelistCodeTransferPlan_ModuleComponentNames() As BetterArray
    EnsureArray this.standardModules, "Standard module list"
    Set ILinelistCodeTransferPlan_ModuleComponentNames = this.standardModules
End Function

Private Function ILinelistCodeTransferPlan_FormNames() As BetterArray
    EnsureArray this.formNames, "Form list"
    Set ILinelistCodeTransferPlan_FormNames = this.formNames
End Function

Private Property Get ILinelistCodeTransferPlan_WorkbookModuleName() As String
    ILinelistCodeTransferPlan_WorkbookModuleName = this.workbookModuleName
End Property
