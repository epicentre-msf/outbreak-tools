Attribute VB_Name = "PrepareSetup"

Option Explicit

'@Folder("Initializations")

'This module prepares the disease setup for usage and creates required elements.

Private dropArray As BetterArray
Private drop As IDropdownLists
Private wb As Workbook
Private devsh As Worksheet

Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub Initialize()
    Dim dropsh As Worksheet
    Set wb = ThisWorkbook
    Set devsh = wb.Worksheets("Dev")
    Set dropsh = wb.Worksheets("__dropdowns")
    'Initilialize the dropdown array and list
    Set dropArray = New BetterArray
    Set drop = DropdownLists.Create(dropsh)
End Sub


'Function to add Elements to the dropdown list
Private Sub AddElements(ByVal dropdownName As String, ParamArray els() As Variant)
    Dim nbEls As Integer
    
    '@Ignore DefaultMemberRequired
    For nbEls = 0 To UBound(els())
        dropArray.Push els(nbEls)
    Next

    drop.Add dropArray, dropdownName
    dropArray.Clear
End Sub


Private Sub CreateDropdowns()
    AddElements "__yes_no", "yes", "no"
    'initialize the diseases list with worksheets names not to put in disease
    AddElements "__diseases_list", "Choices", "Variables", "Translations", _
                 "Dev", "__pass", "__updated", "__compRep", "__ribbonTranslation", _
                 "__dropdowns"

    AddElements "__configSheets", "Dev", "__pass", "__updated", "__compRep", "__ribbonTranslation", "__dropdowns"
    AddElements "__languages", vbNullString
    AddElements "__file_languages", vbNullString
    AddElements "__var_status", "mandatory", "optional, visible", "optional, hidden"
End Sub

Private Sub AddLanguages()
    Dim Lo As ListObject
    Dim langList As BetterArray

    Set Lo = wb.Worksheets("Translations").ListObjects(1)
    Set langList = New BetterArray

    langList.FromExcelRange Lo.HeaderRowRange
    drop.Update UpdateData:=langList, listName:="__languages"
End Sub


'@Description("Configure the setup for codes")
'@EntryPoint
Public Sub ConfigureSetup()
Attribute ConfigureSetup.VB_Description = "Configure the setup for codes"
    
    'For production mode setup, exit without configuring
    Initialize

    On Error Resume Next
    If devsh.Range("RNG_InProduction").Value = "yes" Then Exit Sub
    On Error GoTo 0

    BusyApp
    CreateDropdowns
    AddLanguages
    EventsRibbon.UpdateWatchedValues
    TransferCodes
    MsgBox "Done!"
    NotBusyApp
End Sub


'Transfercodes to Worksheet

Private Sub TransferCodes()

   Dim objectsList As BetterArray                'List of sheets where to transfer the code
   Dim counter As Long
   Dim sheetName As String

   Set objectsList = New BetterArray

  'Workbooklevel is just a tag to import changes at the workbook level.
   objectsList.Push "__WorkbookLevel", "Variables", "Choices"

   For counter = objectsList.LowerBound To objectsList.UpperBound
        sheetName = objectsList.Item(counter)
        Misc.TransferCodeWksh sheetName
   Next
End Sub



'@Description("Prepare the disease setup for production")
'@EntryPoint
Public Sub PrepareForProd()
Attribute PrepareForProd.VB_Description = "Prepare the disease setup for production"
    Dim wb As Workbook
    Dim pass As IPasswords
    Dim pwd As String
    Dim sh As Worksheet

    Set wb = ThisWorkbook

    On Error Resume Next
    If (wb.Worksheets("Dev").Range("RNG_InProduction").Value = "yes") Then
        Exit Sub
    End If
    On Error GoTo 0

    BusyApp

    'First write the password to the password sheet
    pwd = wb.Worksheets("Dev").Range("RNG_DevPasswd").Value
    wb.Worksheets("__pass").Range("RNG_DebuggingPassword").Value = pwd

    'Protect the worksheets
    Set sh = wb.Worksheets("__pass")
    Set pass = Passwords.Create(sh)
    'As Dictionary
    pass.Protect "Variables"
    'Choices
    pass.Protect "Choices"
    'Translations
    pass.Protect "Translations", True, True
    'Hide some worksheeets
    pass.UnProtectWkb wb

    wb.Worksheets("__updated").Visible = xlSheetHidden
    wb.Worksheets("__pass").Visible = xlSheetHidden
    wb.Worksheets("__dropdowns").Visible = xlSheetHidden
    wb.Worksheets("Dev").Range("RNG_InProduction").Value = "yes"
    wb.Worksheets("Dev").Visible = xlSheetHidden

    'Protect the workbook
    pass.ProtectWkb wb

    'Protect the project
    NotBusyApp
End Sub
