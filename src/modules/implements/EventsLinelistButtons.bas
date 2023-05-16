Attribute VB_Name = "EventsLinelistButtons"
Attribute VB_Description = "Events associated to eventual buttons in the Linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated with click on buttons in the linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"

Private showHideObject As ILLShowHide
Private tradsform As ITranslation   'Translation of forms
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
    Set tradsform = lltrads.TransObject(TranslationOfForms)
End Sub

'@Description("Callback for click on show/hide in a linelist worksheet on a button")
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

'@Description("Callback for click on the list of showhide")
'@EntryPoint
Public Sub ClickListShowHide(ByVal index As Long)
    showHideObject.UpdateVisibilityStatus index
End Sub

'@Description("Callback for clik on differents show hide options on a button")
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal index As Long)
    showHideObject.ShowHideLogic index
End Sub


'@Description("Callback for click on the Print Button")
'@EntryPoint
Public Sub ClickOpenPrint()
    Const PRINTPREFIX As String = "print_"

    Dim sh As Worksheet
    Dim printsh As Worksheet
    Dim wb As Workbook
    Dim sheetTag As String

    Set wb = ThisWorkbook

    Set pass = LLPasswords.Create(wb.Worksheets(PASSSHEET))
    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" Then Exit Sub

    Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
    'UnProtect current workbook
    pass.UnprotectWkb wb
    'Unhide the linelist Print
    printsh.Visible = xlSheetVisible
    printsh.Activate
    pass.ProtectWkb wb
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal index As Long)
    showHideObject.ChangeColWidth index
End Sub