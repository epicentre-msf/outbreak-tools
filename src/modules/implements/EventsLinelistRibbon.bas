Attribute VB_Name = "EventsLinelistRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated with the Ribbon Menu in the linelist")

Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private showHideObject As ILLShowHide
Private tradsform As ITranslation   'Translation of forms
Private tradsmess As ITranslation   'Translation of messages

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet


    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)

    Set tradsmess = lltrads.TransObject()
    Set tradsform = lltrads.TransObject(TranslationOfForms)
End Sub

'Callback for click on show/hide in a linelist worksheet
'@EntryPoint
Public Sub ClickShowHide()
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetTag As String
    Dim fobject As Object 'either translation of shapes or translation of formula

    Set sh = ActiveSheet
    'Test the sheet type to be sure it is a HList or a HList Print,
    'and exit if not
    sheetTag = sh.Cells(1, 3).Value
    If sheetTag <> "HList" And sheetTag <> "HList Print" Then Exit Sub

    'initialize the translations of forms and messages
    InitializeTrads

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets(DICTSHEET), 1, 1)

    'This is the private show hide object, used in future subs.
    Set showHideObject = LLShowHide.Create(tradsmess, dict, sh)

    'Load elements to the current form
    showHideObject.Load tradsform
End Sub

'Callback for click on the list of showhide
'@EntryPoint
Public Sub ClickListShowHide(ByVal index As Long)
    showHideObject.UpdateVisibilityStatus index
End Sub

'Callback for clik on differents show hide options
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal index As Long)
    showHideObject.ShowHideLogic index
End Sub
