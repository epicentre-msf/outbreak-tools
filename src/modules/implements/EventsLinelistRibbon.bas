Attribute VB_Name = "EventsLinelistRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated with the Ribbon Menu in the linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"

Private tradrib As ITranslation   'Translation of forms
Private tradsmess As ITranslation   'Translation of messages
Private pass As ILLPasswords
Private ribbonUI As IRibbonUI

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

'@Description("Callback for customUI.onLoad")
'@EntryPoint
Public Sub ribbonLinelistLoaded(ribbon As IRibbonUI)
     Set ribbonUI = ribbon
End Sub

'@Description("Callback for adminTab getLabel")
'@EntryPoint
Public Sub getLLLang(control As IRibbonControl, ByRef returnedVal)
    Dim codeId As String
    codeId = control.Id
    returnedVal = tradrib.TranslationMsg(codeId)
End Sub

'@Description("Callback for btnAdvanced onAction")
'@EntryPoint
Public Sub clickRibbonAdvanced(control As IRibbonControl)
End Sub

'@Description("Callback for btnExport onAction")
'@EntryPoint
Public Sub clickRibbonExport(control As IRibbonControl)
End Sub

'@Description("Callback for btnDebug onAction")
'@EntryPoint
Public Sub clickRibbonDegug(control As IRibbonControl)
End Sub

'@Description("Callback for btnShowHideVar onAction")
'@EntryPoint
Public Sub clickRibbonShowHideVar(control As IRibbonControl)
End Sub

'@Description("Callback for btnShowHideSec onAction")
'@EntryPoint
Public Sub clickRibbonShowHideSec(control As IRibbonControl)
End Sub

'@Description("Callback for btnAddRows onAction")
'@EntryPoint
Public Sub clickRibbonAddRows(control As IRibbonControl)
End Sub

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonResize(control As IRibbonControl)
End Sub

'@Description("Callback for btnRemFilt onAction")
'@EntryPoint
Public Sub clickRibbonRemoveFilter(control As IRibbonControl)
End Sub

'@Description("Callback for btnCustomFilt onAction")
'@EntryPoint
Public Sub clickRibbonCustomFilter(control As IRibbonControl)
End Sub

'@Description("Callback for btnOpenPrint onAction")
'@EntryPoint
Public Sub clickRibbonOpenPrint(control As IRibbonControl)
End Sub

'@Description("Callback for btnClosePrint onAction")
'@EntryPoint
Public Sub clickRibbonClosePrint(control As IRibbonControl)
End Sub

'@Description("Callback for btnRotateHead onAction")
'@EntryPoint
Public Sub clickRibbonRotateAll(control As IRibbonControl)
End Sub

'@Description("Callback for btnRowHeight onAction")
'@EntryPoint
Public Sub clickRibbonRowHeight(control As IRibbonControl)
End Sub

'@Description("Callback for btnCalc onAction")
'@EntryPoint
Public Sub clickRibbonCalculate(control As IRibbonControl)
End Sub

'@Description("Callback for btnApplyFilt onAction")
'@EntryPoint
Public Sub clickRibbonApplyFilt(control As IRibbonControl)
End Sub

