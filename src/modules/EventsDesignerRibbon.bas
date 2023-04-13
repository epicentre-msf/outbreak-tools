Attribute VB_Name = "EventsDesignerRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the designer"
Option Explicit
'@Folder("Designer Events")
'@ModuleDescription("Events associated to the Ribbon Menu in the designer")


'@Description("Callback for getLabel (Depending on the language)")
'@EntryPoint
Public Sub LangLabel(control As IRibbonControl, ByRef returnedVal)
Attribute LangLabel.VB_Description = "Callback for getLabel (Depending on the language)"
End Sub

'@Description("Callback for btnDelGeo onAction: Delete the geobase")
'@EntryPoint
Public Sub clickDelGeo(control As IRibbonControl)
Attribute clickDelGeo.VB_Description = "Callback for btnDelGeo onAction: Delete the geobase"
End Sub

'@Description("Callback for btnClear onAction": Clear the entries)
'@EntryPoint
Public Sub clickClearEnt(control As IRibbonControl)
End Sub

'@Description("Callback for btnTransAdd onAction: Import Linelist translations")
'@EntryPoint
Public Sub clickImpTrans(control As IRibbonControl)
Attribute clickImpTrans.VB_Description = "Callback for btnTransAdd onAction: Import Linelist translations"
End Sub

'@Description("Callback for langDrop onAction: Change the language of the designer")
'@EntryPoint
Public Sub clickLangChange(control As IRibbonControl, id As String, Index As Integer)
Attribute clickLangChange.VB_Description = "Callback for langDrop onAction: Change the language of the designer"
End Sub

'@Description("Callback for btnOpen onAction: Open another linelist file")
'@EntryPoint
Public Sub clickOpen(control As IRibbonControl)
Attribute clickOpen.VB_Description = "Callback for btnOpen onAction: Open another linelist file"
End Sub
