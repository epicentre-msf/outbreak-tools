Attribute VB_Name = "EventsLinelistRibbon"
Attribute VB_Description = "Events associated with the Ribbon Menu in the linelist"
Option Explicit
Option Private Module

'@IgnoreModule ParameterNotUsed
'@Folder("Linelist Events")
'@ModuleDescription("Events associated with the Ribbon Menu in the linelist")

Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

Private tradrib As ITranslation   'Translation of forms

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet

    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradrib = lltrads.TransObject(TranslationOfRibbon)

End Sub

'@Description("Callback for adminTab getLabel")
'@EntryPoint
Public Sub getLLLang(ByRef Control As IRibbonControl, ByRef returnedVal)
    Attribute getLLLang.VB_Description = "Callback for adminTab getLabel"

    Dim codeId As String
    InitializeTrads
    codeId = Control.ID
    returnedVal = tradrib.TranslatedValue(codeId)
End Sub

'@Description("Callback for btnAdvanced onAction")
'@EntryPoint
Public Sub clickRibbonAdvanced(ByRef Control As IRibbonControl)
    Attribute clickRibbonAdvanced.VB_Description = "Callback for btnAdvanced onAction"

    'Call the clickAdvanced from buttons
    ClickAdvanced
End Sub

'@Description("Callback for btnExport onAction")
'@EntryPoint
Public Sub clickRibbonExport(ByRef Control As IRibbonControl)
    Attribute clickRibbonExport.VB_Description = "Callback for btnExport onAction"
    'call the clickExport from buttons
    ClickExport
End Sub

'@Description("Callback for btnDebug onAction")
'@EntryPoint
Public Sub clickRibbonDegug(ByRef Control As IRibbonControl)
    Attribute clickRibbonDegug.VB_Description = "Callback for btnDebug onAction"
End Sub

'@Description("Callback for btnShowHideVar onAction")
'@EntryPoint
Public Sub clickRibbonShowHideVar(ByRef Control As IRibbonControl)
    Attribute clickRibbonShowHideVar.VB_Description = "Callback for btnShowHideVar onAction"
    ClickShowHide
End Sub

'@Description("Callback for btnShowHideSec onAction")
'@EntryPoint
Public Sub clickRibbonShowHideSec(ByRef Control As IRibbonControl)
    Attribute clickRibbonShowHideSec.VB_Description = "Callback for btnShowHideSec onAction"
End Sub

'@Description("Callback for btnAddRows onAction")
'@EntryPoint
Public Sub clickRibbonAddRows(ByRef Control As IRibbonControl)
    Attribute clickRibbonAddRows.VB_Description = "Callback for btnAddRows onAction"
    ClickAddRows
End Sub

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonResize(ByRef Control As IRibbonControl)
    Attribute clickRibbonResize.VB_Description = "Callback for btnResize onAction"
    ClickResize
End Sub

'@Description("Callback for btnRemFilt onAction")
'@EntryPoint
Public Sub clickRibbonRemoveFilter(ByRef Control As IRibbonControl)
    Attribute clickRibbonRemoveFilter.VB_Description = "Callback for btnRemFilt onAction"
    ClickRemoveFilters
End Sub

'@Description("Callback for btnOpenPrint onAction")
'@EntryPoint
Public Sub clickRibbonOpenPrint(ByRef Control As IRibbonControl)
    Attribute clickRibbonOpenPrint.VB_Description = "Callback for btnOpenPrint onAction"
    ClickOpenPrint
End Sub

'@Description("Callback for btnOpenForm onAction")
'@EntryPoint
Public Sub clickRibbonOpenCRF(ByRef Control As IRibbonControl)
    Attribute clickRibbonOpenCRF.VB_Description = "Callback for btnOpenForm onAction"
    ClickOpenCRF
End Sub

'@Description("Callback for btnClosePrint onAction")
'@EntryPoint
Public Sub clickRibbonClosePrint(ByRef Control As IRibbonControl)
    Attribute clickRibbonClosePrint.VB_Description = "Callback for btnClosePrint onAction"
    ClickClosePrint
End Sub

'@Description("Callback for btnRotateHead onAction")
'@EntryPoint
Public Sub clickRibbonRotateAll(ByRef Control As IRibbonControl)
    Attribute clickRibbonRotateAll.VB_Description = "Callback for btnRotateHead onAction"
    ClickRotateAll
End Sub

'@Description("Callback for btnRowHeight onAction")
'@EntryPoint
Public Sub clickRibbonRowHeight(ByRef Control As IRibbonControl)
    Attribute clickRibbonRowHeight.VB_Description = "Callback for btnRowHeight onAction"
    ClickRowHeight
End Sub

'@Description("Callback for btnCalc onAction")
'@EntryPoint
Public Sub clickRibbonCalculate(ByRef Control As IRibbonControl)
    Attribute clickRibbonCalculate.VB_Description = "Callback for btnCalc onAction"
    ClickCalculate
End Sub

'@Description("Callback for btnApplyFilt onAction")
'@EntryPoint
Public Sub clickRibbonApplyFilt(ByRef Control As IRibbonControl)
    Attribute clickRibbonApplyFilt.VB_Description = "Callback for btnApplyFilt onAction"
End Sub

'@Description("Callback for btnGeo onAction")
'@EntryPoint
Public Sub clickRibbonGeo(ByRef Control As IRibbonControl)
    Attribute clickRibbonGeo.VB_Description = "Callback for btnGeo onAction"
    ClickGeoApp
End Sub

'@Description("Callback for btnPrintLL onAction")
'@EntryPoint
Public Sub clickRibbonPrintLL(ByRef Control As IRibbonControl)
    Attribute clickRibbonPrintLL.VB_Description = "Callback for btnPrintLL onAction"
    ClickPrintLL
End Sub

'@Description("Callback for btnOpenLab onAction")
'@EntryPoint
Public Sub clickRibbonOpenVarLab(ByRef Control As IRibbonControl)
    Attribute clickRibbonOpenVarLab.VB_Description = "Callback for btnOpenLab onAction"
    ClickOpenVarLab
End Sub

'@Description("Callback for btnSortTab on Action")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef Control As IRibbonControl)
    Attribute clickRibbonSortTable.VB_Description = "Callback for btnSortTab on Action"
    ClickSortTable
End Sub

'@Description("Callback for btnExpAna on Action")
'@EntryPoint
Public Sub clickRibbonExportAnalysis(ByRef Control As IRibbonControl)
    Attribute clickRibbonExportAnalysis.VB_Description = "Callback for btnExpAna on Action"
    ClickExportAnalysis
End Sub

'@Description("Callback for btnImport On Action")
'@EntryPoint
Public Sub clickRibbonImport(ByRef Control As IRibbonControl)
    Attribute clickRibbonImport.VB_Description = "Callback for btnImport On Action"
    ClickImportData
End Sub

'@Description("Callback for btnImportGeo On Action")
'@EntryPoint
Public Sub clickRibbonImportGeobase(ByRef Control As IRibbonControl)
    Attribute clickRibbonImportGeobase.VB_Description = "Callback for btnImport On Action"
    ClickImportGeobase
End Sub

'@Description("Callback for btnAutoFit On Action")
'@EntryPoint
Public Sub clickRibbonAutoFit(ByRef Control As IRibbonControl)
    Attribute clickRibbonAutoFit.VB_Description = "Callback for btnAutoFit On Action"
    clickAutoFit
End Sub


'@Description("Callback for btnSetEpiWeek On Action")
'@EntryPoint
Public Sub clickRibbonSetEpiWeek(ByRef Control As IRibbonControl)
 Attribute clickRibbonSetEpiWeek.VB_Description = "Callback for btnSetEpiWeek On Action"
    [F_EpiWeek].Show
End Sub