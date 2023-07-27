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
Public Sub getLLLang(ByRef control As IRibbonControl, ByRef returnedVal)
    Dim codeId As String
    InitializeTrads
    codeId = control.Id
    returnedVal = tradrib.TranslatedValue(codeId)
End Sub

'@Description("Callback for btnAdvanced onAction")
'@EntryPoint
Public Sub clickRibbonAdvanced(ByRef control As IRibbonControl)
    'Call the clickAdvanced from buttons
    ClickAdvanced
End Sub

'@Description("Callback for btnExport onAction")
'@EntryPoint
Public Sub clickRibbonExport(ByRef control As IRibbonControl)
    'call the clickExport from buttons
    ClickExport
End Sub

'@Description("Callback for btnDebug onAction")
'@EntryPoint
Public Sub clickRibbonDegug(ByRef control As IRibbonControl)
End Sub

'@Description("Callback for btnShowHideVar onAction")
'@EntryPoint
Public Sub clickRibbonShowHideVar(ByRef control As IRibbonControl)
    ClickShowHide
End Sub

'@Description("Callback for btnShowHideSec onAction")
'@EntryPoint
Public Sub clickRibbonShowHideSec(ByRef control As IRibbonControl)
End Sub

'@Description("Callback for btnAddRows onAction")
'@EntryPoint
Public Sub clickRibbonAddRows(ByRef control As IRibbonControl)
    ClickAddRows
End Sub

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonResize(ByRef control As IRibbonControl)
    ClickResize
End Sub

'@Description("Callback for btnRemFilt onAction")
'@EntryPoint
Public Sub clickRibbonRemoveFilter(ByRef control As IRibbonControl)
    ClickRemoveFilters
End Sub

'@Description("Callback for btnCustomFilt onAction")
'@EntryPoint
Public Sub clickRibbonCustomFilter(ByRef control As IRibbonControl)
End Sub

'@Description("Callback for btnOpenPrint onAction")
'@EntryPoint
Public Sub clickRibbonOpenPrint(ByRef control As IRibbonControl)
    ClickOpenPrint
End Sub

'@Description("Callback for btnClosePrint onAction")
'@EntryPoint
Public Sub clickRibbonClosePrint(ByRef control As IRibbonControl)
    ClickClosePrint
End Sub

'@Description("Callback for btnRotateHead onAction")
'@EntryPoint
Public Sub clickRibbonRotateAll(ByRef control As IRibbonControl)
    ClickRotateAll
End Sub

'@Description("Callback for btnRowHeight onAction")
'@EntryPoint
Public Sub clickRibbonRowHeight(ByRef control As IRibbonControl)
    ClickRowHeight
End Sub

'@Description("Callback for btnCalc onAction")
'@EntryPoint
Public Sub clickRibbonCalculate(ByRef control As IRibbonControl)
    ClickCalculate
End Sub

'@Description("Callback for btnApplyFilt onAction")
'@EntryPoint
Public Sub clickRibbonApplyFilt(ByRef control As IRibbonControl)
End Sub

'@Description("Callback for btnGeo onAction")
'@EntryPoint
Public Sub clickRibbonGeo(ByRef control As IRibbonControl)
    ClickGeoApp
End Sub

'@Description("Callback for btnPrintLL onAction")
'@EntryPoint
Public Sub clickRibbonPrintLL(ByRef control As IRibbonControl)
    ClickPrintLL
End Sub

'@Description("Callback for btnOpenLab onAction")
'@EntryPoint
Public Sub clickRibbonOpenVarLab(ByRef control As IRibbonControl)
    ClickOpenVarLab
End Sub