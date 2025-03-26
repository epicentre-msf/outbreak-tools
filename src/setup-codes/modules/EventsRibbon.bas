Attribute VB_Name = "EventsRibbon"

Option Explicit

'@Folder("Events")
'@IgnoreModule SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const PASSSHEETNAME As String = "__pass"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const DICTSHEETNAME As String = "Dictionary"

'Private Subs to speed up process
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.EnableAnimations = True
End Sub

'@Description("Resize the listObjects in the current sheet")
'@EntryPoint
Public Sub clickResize(ByRef control As IRibbonControl)
Attribute clickResize.VB_Description = "Resize the listObjects in the current sheet"
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=True
End Sub

'@Description("add rows to listObject")
'@EntryPoint
Public Sub clickAddRows(ByRef control As Office.IRibbonControl)
Attribute clickAddRows.VB_Description = "add rows to listObject"
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    ManageRows sheetName:=sheetName, del:=False
End Sub

'@Description("Callback for editLang onChange: Add a language to translation table")
'@EntryPoint
Public Sub clickAddLang(ByRef control As IRibbonControl, ByRef text As String)
Attribute clickAddLang.VB_Description = "Callback for editLang onChange: Add a language to translation table"

    Dim pass As IPasswords
    Dim trads As ITranslations
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim askFirst As Long

    If text = vbNullString Then Exit Sub 
    BusyApp

    'Ask before proceeding
    askFirst = MsgBox("Do you really want to add language(s) " & _
                      text & " to translations?", _
                      vbYesNo, "Confirm")

    If (askFirst = vbNo) Then Exit Sub

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(TRADSHEETNAME)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translations.Create(sh, TABTRANSLATION)

    pass.UnProtect TRADSHEETNAME
    trads.AddTransLang text
    pass.Protect TRADSHEETNAME, True, True

    MsgBox "Done!"
    NotBusyApp
End Sub

'@Description("Callback for btnTransAdd onAction: Import all words to be translated")
'@EntryPoint
Public Sub clickAddTrans(ByRef control As IRibbonControl)
Attribute clickAddTrans.VB_Description = "Callback for btnTransAdd onAction: Import all words to be translated"
    Dim pass As IPasswords
    Dim trads As ITranslations
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim upsh As Worksheet
    Dim askFirst As Long

    BusyApp
    'Ask before proceeding
    askFirst = MsgBox("Do you want to update the translation sheet?", vbYesNo, "Confirm")

    If (askFirst = vbNo) Then Exit Sub

    Application.Cursor = xlWait
    On Error GoTo errHand

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADSHEETNAME)
    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translations.Create(tradsh, TABTRANSLATION)

    pass.UnProtect TRADSHEETNAME
    On Error Resume Next
    tradsh.ListObjects(1).AutoFilter.ShowAllData
    On Error GoTo 0
    'update all values for translation
    trads.UpdateTrans upsh
    pass.Protect TRADSHEETNAME, True, True

    'Set all updates to no (this sub is in the EventsGlobal module)
    SetAllUpdatedTo "no"
    NotBusyApp
    Application.Cursor = xlDefault
    Exit Sub

errHand:
    Application.Cursor = xlDefault
    MsgBox "An internal error occured, contact the developper", vbCritical + vbOkOnly
    NotBusyApp
End Sub

'@Description("Callback for btnTransUp onAction: Update columns to be translated")
'@EntryPoint
Public Sub clickUpdateTranslate(ByRef control As IRibbonControl)
Attribute clickUpdateTranslate.VB_Description = "Callback for btnTransUp onAction: Update columns to be translated"
    'remove update columns and add new columns to watch
    
    On Error Resume Next
    If (ThisWorkbook.Worksheets("Dev").Range("RNG_InProduction").Value = "yes") Then
        Exit Sub
    End If
    On Error GoTo 0

    BusyApp
    CleanUpdateColumns
    UpdatedWatchedValues
    NotBusyApp
    MsgBox "Done!"
End Sub

'@Description("Callback for btnChk onAction: Check the setup for eventual errors")
'@EntryPoint
Public Sub clickCheck(ByRef control As IRibbonControl)
Attribute clickCheck.VB_Description = "Callback for btnChk onAction: Check the setup for eventual errors"
    BusyApp
    Dim askFirst As Long
    askFirst = MsgBox("Do you really want to check the current setup?", vbYesNo, "Confirmation")
    If askFirst = vbYes Then CheckTheSetup
    ThisWorkbook.Worksheets("__checkRep").Activate
    NotBusyApp
End Sub

'@Description("Callback for btnImp onAction: Import elements from another setup")
'@EntryPoint
Public Sub clickImport(ByRef control As IRibbonControl)
Attribute clickImport.VB_Description = "Callback for btnImp onAction: Import elements from another setup"
    PrepareForm cleanSetup:=False
    [Imports].Show
End Sub

'@Description("Callback for btnClear onAction: clean the setup")
'@EntryPoint
Public Sub clickClearSetup(ByRef control As IRibbonControl)
Attribute clickClearSetup.VB_Description = "Callback for btnClear onAction: clean the setup"
    PrepareForm cleanSetup:=True
    [Imports].Show
End Sub

'@Description("Callback for btnInit onAction: Initialize the setup")
'@EntryPoint
Public Sub clickInitializeSetup(ByRef control As IRibbonControl)
Attribute clickInitializeSetup.VB_Description = "Callback for btnInit onAction: Initialize the setup"
    'Call configure setup from the PrepareSetup module
    ConfigureSetup
End Sub

'@Description("Callback for btnPrep onAction: Prepare the setup for deployment")
'@EntryPoint
Public Sub clickPrepareProd(ByRef control As IRibbonControl)
Attribute clickPrepareProd.VB_Description = "Callback for btnPrep onAction: Prepare the setup for deployment"
    'Call prepare for production from the PrepareSetup module
    PrepareForProd
End Sub

'@Description("Callback for btnFilt onAction: clear all the filters in the current sheet")
'@EntryPoint
Public Sub clickFilters(ByRef control As IRibbonControl)
Attribute clickFilters.VB_Description = "Callback for btnFilt onAction: clear all the filters in the current sheet"

    Dim pass As IPasswords
    Dim shpass As Worksheet
    Dim sh As Worksheet
    Dim sheetName As String
    Dim Lo As ListObject

    BusyApp
    Set sh = ActiveSheet
    sheetName = sh.Name
    Set shpass = ThisWorkbook.Worksheets(PASSSHEETNAME)
    Set pass = Passwords.Create(shpass)
    pass.Unprotect sheetName

    For Each Lo in sh.ListObjects
        If Not Lo.AutoFilter Is Nothing Then
            On Error Resume Next
                Lo.AutoFilter.ShowAllData
            On Error GoTo 0
        End If
    Next
    'Analysis and translations need strict protections
    pass.Protect sheetName,  _
    (sheetName = "Translations" Or sheetName = "Analysis"), _
    (sheetName = "Translations")

    NotBusyApp
End Sub

'===== Auxilliary subs used in the process of adding Rows

'Add or Remove Rows to a table
Public Sub ManageRows(ByVal sheetName As String, _
                      Optional ByVal del As Boolean = False, _ 
                      Optional ByVal allAnalysis As Boolean = False)
    Dim part As Object
    Dim sh As Worksheet
    Dim shpass As Worksheet
    Dim pass As IPasswords
    Dim prevNbExp As Long
    Dim actNbExp As Long


    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    sh.EnableCalculation = False
    Set shpass = ThisWorkbook.Worksheets(PASSSHEETNAME)
    On Error GoTo 0

    If (sh Is Nothing) Or (shpass Is Nothing) Then Exit Sub

    '5 is the start line of the dictionary
    '4 is the start column of the dictionary
    Select Case sheetName
    Case "Dictionary"
        Set part = LLdictionary.Create(sh, 5, 1)
    Case "Choices"
        Set part = LLchoice.Create(sh, 4, 1)
    Case "Analysis"
        If allAnalysis Then sh.Range("RNG_SelectTable").Value = "Add or remove rows of all tables"
        Set part = Analysis.Create(sh)
    Case "Exports"
        Set part = LLExport.Create(sh, 4, 1)
        prevNbExp = part.NumberOfExports()
    End Select

    'Exit if unable to find the corresponding object
    If part Is Nothing Then Exit Sub

    BusyApp
    Set pass = Passwords.Create(shpass)
    BusyApp
    pass.UnProtect sh.Name

    If del Then
        part.RemoveRows
    Else
        part.AddRows
    End If


    If (sh.Name = "Exports") Then 
        actNbExp = part.NumberOfExports()
        ManageDictionaryExport prevNbExp, actNbExp, del
    End If

    BusyApp
    pass.Protect sh.Name, (sh.Name = "Analysis")
    sh.EnableCalculation = True
    NotBusyApp
End Sub

Private Sub ManageDictionaryExport(ByVal prevNbExp As Long, ByVal actNbExp As Long, _ 
                                   Optional ByVal del As Boolean = False)
    Dim dict As ILLdictionary
    Dim counter As Long
    Dim shpass As Worksheet
    Dim shdict As Worksheet
    Dim pass As IPasswords

    Set shpass = ThisWorkbook.Worksheets(PASSSHEETNAME)
    Set shdict = ThisWorkbook.Worksheets(DICTSHEETNAME)
    Set pass = Passwords.Create(shpass)

    BusyApp
    pass.UnProtect "Dictionary"
    Set dict = LLdictionary.Create(shdict, 5, 1)
    
    If del Then
        'Delete the exports columns from the dictionary (previous number > actual number)
        For counter = (actNbExp + 1) To prevNbExp
            dict.RemoveColumn "Export " & counter
        Next 
    Else
        'Add new exports columns to the dictionary (previous number < actual number)
        For counter = (prevNbExp + 1) To actNbExp
            dict.InsertColumn "Export " & counter, After:= "Export " & (counter - 1)
        Next
    End If

    BusyApp
    pass.Protect "Dictionary"
End Sub

Private Sub CleanUpdateColumns()
    'Clear the update sheet
    Dim upsh As Worksheet
    Dim Lo As ListObject
    Dim wb As Workbook
    Dim namesRng As Range
    Dim counter As Long
    Set wb = ThisWorkbook
    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)

    'Unlist all listObjects in the worksheet and delete all names
    For Each Lo In upsh.ListObjects
        Set namesRng = Lo.ListColumns("rngname").Range
        For counter = 1 To namesRng.Rows.Count
            On Error Resume Next
            wb.Names(namesRng.Cells(counter, 1).Value).Delete
            On Error GoTo 0
        Next
        Lo.Unlist
    Next
    upsh.Cells.Clear
End Sub

'Update the translation values
'@EntryPoint
Public Sub UpdatedWatchedValues()
    Dim sh As Worksheet
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String

    Set sheetsList = New BetterArray
    sheetsList.Push "Dictionary", "Choices", "Exports", "Analysis"
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        Set sh = ThisWorkbook.Worksheets(sheetName)
        'Write update status on each sheet
        writeUpdateStatus sh
    Next
End Sub

'Update status of columns to watch
Private Sub writeUpdateStatus(sh As Worksheet)
    Dim upsh As Worksheet
    Dim upId As String
    Dim upObj As IUpdatedValues
    Dim Lo As ListObject

    Set upsh = ThisWorkbook.Worksheets(UPDATEDSHEETNAME)
    upId = LCase(Left(sh.Name, 4))
    For Each Lo In sh.ListObjects
        If sh.Name = "Analysis" Then _
        upId = LCase(Replace(Lo.Name, "Tab_", vbNullString))
        Set upObj = UpdatedValues.Create(upsh, upId)
        upObj.AddColumns Lo
    Next
End Sub

'Prepare the Import Form
Private Sub PrepareForm(Optional ByVal cleanSetup As Boolean = False)
    If cleanSetup Then
        [Imports].LoadButton.Visible = False
        [Imports].LabPath.Visible = False
        [Imports].InfoChoice.Caption = "Select what to Clear"
        [Imports].DictionaryCheck.Caption = "Clear Dictionary"
        [Imports].ChoiceCheck.Caption = "Clear Choices"
        [Imports].ExportsCheck.Caption = "Clear Exports"
        [Imports].AnalysisCheck.Caption = "Clear Analysis"
        [Imports].TranslationsCheck.Caption = "Clear Translation"
        [Imports].ConformityCheck.Visible = False
        [Imports].DoButton.Caption = "Clear"

        'Resize and change position of elements
        [Imports].Height = 400
        [Imports].InfoChoice.Top = 20
        [Imports].DictionaryCheck.Top = 50
        [Imports].ChoiceCheck.Top = 80
        [Imports].ExportsCheck.Top = 110
        [Imports].AnalysisCheck.Top = 140
        [Imports].TranslationsCheck.Top = 170
        [Imports].LabProgress.Top = 200
        [Imports].DoButton.Top = 270
        [Imports].Quit.Top = 310
    Else
        [Imports].InfoChoice.Caption = "Select what to Import"
        [Imports].DictionaryCheck.Caption = "Import Dictionary"
        [Imports].ChoiceCheck.Caption = "Import Choices"
        [Imports].ExportsCheck.Caption = "Import Exports"
        [Imports].AnalysisCheck.Caption = "Import Analysis"
        [Imports].TranslationsCheck.Caption = "Import Translation"
        [Imports].ConformityCheck.Visible = True
        [Imports].LoadButton.Visible = True
        [Imports].LabPath.Visible = True
        [Imports].DoButton.Caption = "Import"

        'resize the worksheet and position of elements
        [Imports].Height = 500
        [Imports].LoadButton.Top = 10
        [Imports].LabPath.Top = 55
        [Imports].InfoChoice.Top = 135
        [Imports].DictionaryCheck.Top = 170
        [Imports].ChoiceCheck.Top = 200
        [Imports].ExportsCheck.Top = 230
        [Imports].AnalysisCheck.Top = 260
        [Imports].TranslationsCheck.Top = 290
        [Imports].DoButton.Top = 350
        [Imports].LabProgress.Top = 390
        [Imports].Quit.Top = 440
    End If
End Sub

'@Description("Callback for btnImpExp onAction: Import elements from a disease worksheet")
'@EntryPoint
Public Sub clickImportDisease(ByRef control As IRibbonControl)
Attribute clickImportDisease.VB_Description = "Callback for btnImpExp onAction: Import elements from a disease worksheet"
End Sub