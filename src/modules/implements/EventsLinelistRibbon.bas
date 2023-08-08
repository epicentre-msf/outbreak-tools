Attribute VB_Name = "EventsLinelistRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the linelist"
Option Explicit
Option Private Module

'@IgnoreModule ParameterNotUsed
'@Folder("Linelist Events")
'@ModuleDescription("Events associated with the Ribbon Menu in the linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"

Private tradrib As ITranslation   'Translation of forms
Private tradsmess As ITranslation   'Translation of messages
Private pass As ILLPasswords

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet


    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)

    Set tradsmess = lltrads.TransObject()
    Set tradrib = lltrads.TransObject(TranslationOfRibbon)
End Sub

'@Description("Callback for adminTab getLabel")
'@EntryPoint
Public Sub getLLLang(ByRef Control As IRibbonControl, ByRef returnedVal)
    Dim codeId As String
    InitializeTrads
    codeId = Control.ID
    returnedVal = tradrib.TranslatedValue(codeId)
End Sub

'@Description("Callback for btnAdvanced onAction")
'@EntryPoint
Public Sub clickRibbonAdvanced(ByRef Control As IRibbonControl)
    'Call the clickAdvanced from buttons
    ClickAdvanced
End Sub

'@Description("Callback for btnExport onAction")
'@EntryPoint
Public Sub clickRibbonExport(ByRef Control As IRibbonControl)
    'call the clickExport from buttons
    ClickExport
End Sub

'@Description("Callback for btnDebug onAction")
'@EntryPoint
Public Sub clickRibbonDegug(ByRef Control As IRibbonControl)
End Sub

'@Description("Callback for btnShowHideVar onAction")
'@EntryPoint
Public Sub clickRibbonShowHideVar(ByRef Control As IRibbonControl)
    ClickShowHide
End Sub

'@Description("Callback for btnShowHideSec onAction")
'@EntryPoint
Public Sub clickRibbonShowHideSec(ByRef Control As IRibbonControl)
End Sub

'@Description("Callback for btnAddRows onAction")
'@EntryPoint
Public Sub clickRibbonAddRows(ByRef Control As IRibbonControl)
    ClickAddRows
End Sub

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonResize(ByRef Control As IRibbonControl)
    ClickResize
End Sub

'@Description("Callback for btnRemFilt onAction")
'@EntryPoint
Public Sub clickRibbonRemoveFilter(ByRef Control As IRibbonControl)
    ClickRemoveFilters
End Sub

'@Description("Callback for btnCustomFilt onAction")
'@EntryPoint
Public Sub clickRibbonCustomFilter(ByRef Control As IRibbonControl)
End Sub

'@Description("Callback for btnOpenPrint onAction")
'@EntryPoint
Public Sub clickRibbonOpenPrint(ByRef Control As IRibbonControl)
    ClickOpenPrint
End Sub

'@Description("Callback for btnClosePrint onAction")
'@EntryPoint
Public Sub clickRibbonClosePrint(ByRef Control As IRibbonControl)
    ClickClosePrint
End Sub

'@Description("Callback for btnRotateHead onAction")
'@EntryPoint
Public Sub clickRibbonRotateAll(ByRef Control As IRibbonControl)
    ClickRotateAll
End Sub

'@Description("Callback for btnRowHeight onAction")
'@EntryPoint
Public Sub clickRibbonRowHeight(ByRef Control As IRibbonControl)
    ClickRowHeight
End Sub

'@Description("Callback for btnCalc onAction")
'@EntryPoint
Public Sub clickRibbonCalculate(ByRef Control As IRibbonControl)
    ClickCalculate
End Sub

'@Description("Callback for btnApplyFilt onAction")
'@EntryPoint
Public Sub clickRibbonApplyFilt(ByRef Control As IRibbonControl)
End Sub

'@Description("Callback for btnGeo onAction")
'@EntryPoint
Public Sub clickRibbonGeo(ByRef Control As IRibbonControl)
    ClickGeoApp
End Sub

'@Description("Callback for btnPrintLL onAction")
'@EntryPoint
Public Sub clickRibbonPrintLL(ByRef Control As IRibbonControl)
    ClickPrintLL
End Sub

'@Description("Callback for btnOpenLab onAction")
'@EntryPoint
Public Sub clickRibbonOpenVarLab(ByRef Control As IRibbonControl)
    ClickOpenVarLab
End Sub

'@Description("Callback for btnSortTab on Action")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef Control As IRibbonControl)
    ClickSortTable
End Sub

'@Description("Callback for btnExpAna on Action")
'@EntryPoint
Public Sub clickRibbonExportAnalysis(ByRef Control As IRibbonControl)
    ClickExportAnalysis
End Sub
