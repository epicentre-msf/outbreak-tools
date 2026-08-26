Attribute VB_Name = "EventsLinelistRibbon"
Attribute VB_Description = "Events associated with the Ribbon Menu in the linelist"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed
'@Folder("Linelist Events")
'@ModuleDescription("Events associated with the Ribbon Menu in the linelist")
'@depends LinelistEventsManager, EventLinelist, LLTranslation, TranslationObject

'@description
'NO Option Private Module HERE, AND THAT IS DELIBERATE.
'
'Every procedure below is named as a callback by a string in customUI.xml, and
'the ribbon resolves those by NAME rather than by a compiled reference.
'Option Private Module takes a module's Public members out of exactly that kind
'of lookup, and this module was the ONLY ribbon module in the project carrying
'it -- SetupRibbon, EventsDesignerCore, EventsDesignerAdvanced,
'EventsDesignerMulti, EventsMasterSetupRibbon and RibbonDev all go without, and
'so did the old stub baked into the ribbon template back when the tab worked.
'
'It was removed while chasing a report that a ribbon button ran the wrong
'callback. That report is NOT proven to come from here -- the delivered project
'also failed to compile at the time, which accounts for it on its own -- so
'treat this as one candidate removed rather than a fault fixed. What is certain
'is that the odd one out was this module, and the cost of matching its six
'siblings is that the callbacks now show up in the macro list.

Private tradrib As TranslationObject   'Translation of ribbon labels

'Initialize translation of ribbon labels
'
'The ribbon table is read once for the life of the workbook. The tab carries 25
'getLabel bindings and every one of them lands here, so the module variable used
'to be overwritten 25 times while the tab was first drawn. A generated linelist
'carries one language, fixed when it was made, so the object below stays right.
'
'The translation helper is the one EventLinelist holds. Building a second one
'here re-validated all five translation tables on every ribbon callback.
Private Sub InitializeTrads()
    Dim lltrads As LLTranslation
    Dim linelistEvents As EventLinelist

    If Not tradrib Is Nothing Then Exit Sub

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set lltrads = linelistEvents.Translation
    If lltrads Is Nothing Then Exit Sub

    Set tradrib = lltrads.TransObject(TranslationOfRibbon)
End Sub

'@Description("Callback for adminTab getLabel")
'@EntryPoint
Public Sub getLLLang(ByRef ribbonControl As IRibbonControl, ByRef returnedVal)
    Attribute getLLLang.VB_Description = "Callback for adminTab getLabel"

    Dim codeId As String
    InitializeTrads
    If tradrib Is Nothing Then Exit Sub

    codeId = ribbonControl.ID
    returnedVal = tradrib.TranslatedValue(codeId)
End Sub

'@Description("Callback for btnAdvanced onAction")
'@EntryPoint
Public Sub clickRibbonAdvanced(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonAdvanced.VB_Description = "Callback for btnAdvanced onAction"

    'Call the clickAdvanced from buttons
    ClickAdvanced
End Sub

'@Description("Callback for btnExport onAction")
'@EntryPoint
Public Sub clickRibbonExport(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonExport.VB_Description = "Callback for btnExport onAction"
    'call the clickExport from buttons
    ClickExport
End Sub

'@Description("Callback for btnShowHideVar onAction")
'@EntryPoint
Public Sub clickRibbonShowHideVar(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonShowHideVar.VB_Description = "Callback for btnShowHideVar onAction"
    ClickShowHide
End Sub

'@Description("Callback for btnShowHideSec onAction")
'@EntryPoint
Public Sub clickRibbonShowHideSec(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonShowHideSec.VB_Description = "Callback for btnShowHideSec onAction"
    'Toggle the sections the current selection touches
    ClickShowHideSection
End Sub

'@Description("Callback for btnAddRows onAction")
'@EntryPoint
Public Sub clickRibbonAddRows(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonAddRows.VB_Description = "Callback for btnAddRows onAction"
    ClickAddRows
End Sub

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonResize(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonResize.VB_Description = "Callback for btnResize onAction"
    ClickResize
End Sub

'@Description("Callback for btnRemFilt onAction")
'@EntryPoint
Public Sub clickRibbonRemoveFilter(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonRemoveFilter.VB_Description = "Callback for btnRemFilt onAction"
    ClickRemoveFilters
End Sub

'@Description("Callback for btnOpenPrint onAction")
'@EntryPoint
Public Sub clickRibbonOpenPrint(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonOpenPrint.VB_Description = "Callback for btnOpenPrint onAction"
    ClickOpenPrint
End Sub

'@Description("Callback for btnOpenForm onAction")
'@EntryPoint
Public Sub clickRibbonOpenCRF(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonOpenCRF.VB_Description = "Callback for btnOpenForm onAction"
    ClickOpenCRF
End Sub

'@Description("Callback for btnClosePrint onAction")
'@EntryPoint
Public Sub clickRibbonClosePrint(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonClosePrint.VB_Description = "Callback for btnClosePrint onAction"
    ClickCloseSheet
End Sub

'@Description("Callback for btnRotateHead onAction")
'@EntryPoint
Public Sub clickRibbonRotateAll(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonRotateAll.VB_Description = "Callback for btnRotateHead onAction"
    ClickRotateAll
End Sub

'@Description("Callback for btnRowHeight onAction")
'@EntryPoint
Public Sub clickRibbonRowHeight(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonRowHeight.VB_Description = "Callback for btnRowHeight onAction"
    ClickRowHeight
End Sub

'@Description("Callback for btnCalc onAction")
'@EntryPoint
Public Sub clickRibbonCalculate(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonCalculate.VB_Description = "Callback for btnCalc onAction"
    ClickCalculate
End Sub

'@Description("Callback for btnGeo onAction")
'@EntryPoint
Public Sub clickRibbonGeo(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonGeo.VB_Description = "Callback for btnGeo onAction"
    ClickGeoApp
End Sub

'@Description("Callback for btnPrintLL onAction")
'@EntryPoint
Public Sub clickRibbonPrintLL(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonPrintLL.VB_Description = "Callback for btnPrintLL onAction"
    ClickPrintLL
End Sub

'@Description("Callback for btnOpenLab onAction")
'@EntryPoint
Public Sub clickRibbonOpenVarLab(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonOpenVarLab.VB_Description = "Callback for btnOpenLab onAction"
    ClickOpenVarLab
End Sub

'@Description("Callback for btnSortTab on Action")
'@EntryPoint
Public Sub clickRibbonSortTable(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonSortTable.VB_Description = "Callback for btnSortTab on Action"
    ClickSortTable
End Sub

'@Description("Callback for btnExpAna on Action")
'@EntryPoint
Public Sub clickRibbonExportAnalysis(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonExportAnalysis.VB_Description = "Callback for btnExpAna on Action"
    ClickExportAnalysis
End Sub

'@Description("Callback for btnImport On Action")
'@EntryPoint
Public Sub clickRibbonImport(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonImport.VB_Description = "Callback for btnImport On Action"
    ClickImportData
End Sub

'@Description("Callback for btnImportGeo On Action")
'@EntryPoint
Public Sub clickRibbonImportGeobase(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonImportGeobase.VB_Description = "Callback for btnImport On Action"
    ClickImportGeobase
End Sub

'@Description("Callback for btnAutoFit On Action")
'@EntryPoint
Public Sub clickRibbonAutoFit(ByRef ribbonControl As IRibbonControl)
    Attribute clickRibbonAutoFit.VB_Description = "Callback for btnAutoFit On Action"
    clickAutoFit
End Sub


'@Description("Callback for btnSetEpiWeek On Action")
'@EntryPoint
Public Sub clickRibbonSetEpiWeek(ByRef ribbonControl As IRibbonControl)
 Attribute clickRibbonSetEpiWeek.VB_Description = "Callback for btnSetEpiWeek On Action"
    F_EpiWeek.ShowDefaultEpiWeek
End Sub
