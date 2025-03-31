
Attribute VB_Name = "EventsRibbon"
Option Explicit

'@Folder("Events")
'@IgnoreModule SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used

'Private constants for Ribbon Events
Private Const TRADSHEETNAME As String = "Translations"
Private Const TABTRANSLATION As String = "Tab_Translations"
Private Const PASSSHEETNAME As String = "__pass"
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const TRADTABLE As String = "TabTransId"
Private Const TRADTABLESHEET As String = "__ribbonTranslation"
Private Const RNG_FileLang As String = "RNG_FileLang"

'All the ribbon object
Private ribbonUI As IRibbonUI

'Private Subs to speed up process
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.EnableAnimations = True
    Application.Cursor = xlDefault
End Sub

'@Description("Resize the listObjects in the current sheet")
Private Sub ResizeLo(Optional ByVal removeRows As Boolean = False, Optional ByVal totalCount As Long = 0)
Attribute clickAddRows.VB_Description = "Resize the listObjects in the current sheet"
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim pass As IPasswords

    On Error GoTo errHand

    Set wb = ThisWorkbook
    Set sh = ActiveSheet
    If sh.Name = TRADSHEETNAME Then Exit Sub

    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))

    'Speed up application
    BusyApp
    pass.Unprotect sh

    For Each Lo in sh.ListObjects
        Set csTab = CustomTable.Create(Lo)
        If removeRows Then
            csTab.RemoveRows totalCount:=totalCount
        Else
            csTab.AddRows
        End If
    Next

    pass.Protect sh
    NotBusyApp
    Exit Sub

errHand:
    Application.Cursor = xlDefault
    MsgBox "An internal error occured, contact the developper", vbCritical + vbOkOnly
    NotBusyApp
End Sub

'@EntryPoint
'@Description("Callback for btnAdd onAction")
Sub clickAddRows(control As IRibbonControl)
Attribute clickAddRows.VB_Description = "Callback for btnAdd onAction"

    Dim drop As IDropdownLists
    Dim dropsh As Worksheet
    Dim configSheets As BetterArray

    Set dropsh = ThisWorkbook.Worksheets("__dropdowns")

    Set drop = DropdownLists.Create(dropsh)
    Set configSheets = New BetterArray

    Set configSheets = drop.Items("__configSheets")

    If Not configSheets.Includes(ActiveSheet.Name) Then ResizeLo
End Sub

'@EntryPoint
'@Description("Callback for btnRes onAction")
Sub clickResize(control As IRibbonControl)
Attribute clickResize.VB_Description = "Callback for btnRes onAction"

    Dim drop As IDropdownLists
    Dim dropsh As Worksheet
    Dim configSheets As BetterArray
    Dim actsh As Worksheet
    Dim totalCount As Long

    Set dropsh = ThisWorkbook.Worksheets("__dropdowns")
    Set drop = DropdownLists.Create(dropsh)
    Set configSheets = New BetterArray
    Set configSheets = drop.Items("__configSheets")

    If Not configSheets.Includes(ActiveSheet.Name) Then 
        Set actsh = ActiveSheet
        totalCount = IIf(actsh.Cells(2, 4).Value = "DISSHEET", 2, 0)
        ResizeLo removeRows:=True, totalCount:=totalCount
    End If
End Sub

'@Description("Callback for btnFilt onAction: clear all the filters in the current sheet")
'@EntryPoint
Public Sub clickFilters(ByRef control As IRibbonControl)
Attribute clickFilters.VB_Description = "Callback for btnFilt onAction: clear all the filters in the current sheet"

    Dim pass As IPasswords
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim Lo As ListObject

    BusyApp
    Set wb = ThisWorkbook
    Set sh = ActiveSheet
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))

    pass.Unprotect sh
    For Each Lo in sh.ListObjects    
        If Not Lo.AutoFilter Is Nothing Then
            On Error Resume Next
                Lo.AutoFilter.ShowAllData
            On Error GoTo 0
        End If    
    Next

    pass.Protect sh
    NotBusyApp
End Sub

'Translation worksheet =========================================================

'@Description("Callback for editLang onChange: Add a language to translation table")
'@EntryPoint
Public Sub clickAddLang(ByRef control As IRibbonControl, ByRef text As String)
Attribute clickAddLang.VB_Description = "Callback for editLang onChange: Add a language to translation table"

    Dim pass As IPasswords
    Dim trads As ITranslation
    Dim tradchk As ITranslationChunks
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim tradTagsh As Worksheet
    Dim Lo As ListObject
    Dim fileLang As String
    Dim drop As IDropdownLists

    If text = vbNullString Then Exit Sub
    BusyApp

    On Error GoTo errHand

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADSHEETNAME)
    Set tradTagsh = wb.Worksheets(TRADTABLESHEET)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set drop = DropdownLists.Create(wb.Worksheets("__dropdowns"))
    Set tradchk = TranslationChunks.Create(tradsh, TABTRANSLATION, drop)
    Set Lo = tradTagsh.ListObjects(TRADTABLE)
    fileLang = wb.Worksheets(TRADTABLESHEET).Range(RNG_FileLang).Value
  
    'Ask before proceeding
    Set trads = Translation.Create(Lo, fileLang)

    If (MsgBox(trads.TranslatedValue("addLang") & text, _
        vbYesNo, trads.TranslatedValue("askConfirm")) = vbNo) Then Exit Sub

    pass.UnProtect TRADSHEETNAME
    tradchk.AddTransLang text
    pass.Protect TRADSHEETNAME, True, True

    MsgBox trads.TranslatedValue("done")

ErrHand:
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
    UpdateWatchedValues
    NotBusyApp
    MsgBox "Done!"
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
Public Sub UpdateWatchedValues()
    Dim sh As Worksheet
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String

    Set sheetsList = New BetterArray
    sheetsList.Push "Variables", "Choices"
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
    For Each Lo In sh.ListObjects
        upId = LCase(Replace(Lo.Name, "Tab_", vbNullString))
        Set upObj = UpdatedValues.Create(upsh, upId)
        upObj.AddColumns Lo
    Next
End Sub

'@Description("Callback for btnTransAdd onAction: Import all words to be translated")
'@EntryPoint
Public Sub clickAddTrans(ByRef control As IRibbonControl)
Attribute clickAddTrans.VB_Description = "Callback for btnTransAdd onAction: Import all words to be translated"

    Dim pass As IPasswords
    Dim trads As ITranslation
    Dim tradchk As ITranslationChunks
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim tradTagsh As Worksheet
    Dim upsh As Worksheet
    Dim askFirst As Long
    Dim Lo As ListObject
    Dim fileLang As String
    Dim drop As IDropdownLists

    BusyApp

    Application.Cursor = xlWait
    On Error GoTo errHand

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADSHEETNAME)
    Set tradTagsh = wb.Worksheets(TRADTABLESHEET)
    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set drop = DropdownLists.Create(wb.Worksheets("__dropdowns"))
    Set tradchk = TranslationChunks.Create(tradsh, TABTRANSLATION, drop)
    Set Lo = tradTagsh.ListObjects(TRADTABLE)
    fileLang = wb.Worksheets(TRADTABLESHEET).Range(RNG_FileLang).Value
  
    'Ask before proceeding
    Set trads = Translation.Create(Lo, fileLang)

    askFirst = MsgBox(trads.TranslatedValue("askTrans"), vbYesNo, trads.TranslatedValue("askConfirm"))
   
    If (askFirst = vbNo) Then Exit Sub

    pass.UnProtect TRADSHEETNAME
    'update all values for translation
    tradchk.UpdateTrans upsh
    pass.Protect TRADSHEETNAME

    'Set all updates to no (this sub is in the EventsGlobal module)
    EventsGlobal.SetAllUpdatedTo "no"
    NotBusyApp
    Application.Cursor = xlDefault
    Exit Sub

errHand:
    Application.Cursor = xlDefault
    MsgBox "An internal error occured, contact the developper", vbCritical + vbOkOnly
    NotBusyApp
End Sub

'Disease Management subs ----------------------------------------------------------------------

'@Description("Callback for btnAddSheet onAction: Add a new disease Worksheet")
'@EntryPoint
Public Sub clickAddSheet(Control As IRibbonControl)
Attribute clickAddSheet.VB_Description = "Callback for btnAddSheet onAction: Add a new disease Worksheet"
    ManageDiseases.AddDisease
End Sub

'@Description("Callback for btnRemSheet onAction: Remove current disease worksheet")
'@EntryPoint
Public Sub clickRemSheet(Control As IRibbonControl)
Attribute clickRemSheet.VB_Description = "Callback for btnRemSheet onAction: Remove current disease worksheet"
    ManageDiseases.RemoveDisease
End Sub


'@Description("Callback for btnClear onAction: Clear all data in the current disease worksheet")
'@EntryPoint
Public Sub clickClearSheet(Control As IRibbonControl)
Attribute clickClearSheet.VB_Description="Callback for btnClear onAction: Clear all data in the current disease worksheet"
    ManageDiseases.ClearDiseaseSheet
End Sub

'Dealing with outside world ----------------------------------------------------------------------

'@Description("Callback for btnExp onAction: Export the current disease file for setup import")
'@EntryPoint
Public Sub clickExpSheet(Control As IRibbonControl)
Attribute clickExpSheet.VB_Description="Callback for btnExp onAction: Export the current disease file for setup import"
    Exports.ExportToSetup
End Sub


'@Description("Callback for btnComp onAction: Compare two diseases")
'@EntryPoint
Public Sub clickComp(Control As IRibbonControl)
Attribute clickComp.VB_Description="Callback for btnComp onAction: Compare two diseases"  
    Misc.Compare
End Sub

'@Description("Callback for btnImp onAction: Import flat disease file")
'@EntryPoint
Public Sub clickImp(Control As IRibbonControl)
Attribute clickImp.VB_Description="Callback for btnImp onAction: Import flat disease file"
    Exports.ImportFlatFile
End Sub


'@EntryPoint
'@Description("Callback for btnExpMig onAction: Export the current file for Migration")
Public Sub clickExp(Control As IRibbonControl)
Attribute clickExp.VB_Description="Callback for btnExpMig onAction: Export the current file for Migration"
    Exports.ExportForMigration
End Sub

'Changing the ribbon language --------------------------------------------------

'@Description("Callback when the button loaded")
'@EntryPoint
Public Sub ribbonLoaded(ByRef ribbon As IRibbonUI)
Attribute ribbonLoaded.VB_Description = "Callback when the button loaded"
    Set ribbonUI = ribbon
End Sub

'Triggers event to update all the labels by relaunching all the callbacks
Private Sub UpdateLabels()
    ribbonUI.Invalidate
End Sub

'@Description("Callback for getLabel (Depending on the language)")
'@EntryPoint
'@Ignore VariableTypeNotDeclared
Public Sub LangLabel(Control As IRibbonControl, ByRef returnedVal)
Attribute LangLabel.VB_Description = "Callback for getLabel (Depending on the language)"

    Dim trads As ITranslation
    Dim codeId As String
    Dim tradsh As Worksheet
    Dim wb As Workbook
    Dim Lo As ListObject
    Dim fileLang As String

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADTABLESHEET)
    Set Lo = tradsh.ListObjects(TRADTABLE)
    fileLang = tradsh.Range(RNG_FileLang).Value
    Set trads = Translation.Create(Lo, fileLang)
    codeId = Control.ID

    returnedVal = trads.TranslatedValue(codeId)
End Sub

'@Description("Callback for langDrop onAction: Change the language of the designer")
'@EntryPoint
Public Sub clickLangChange(Control As IRibbonControl, langId As String, Index As Integer)
Attribute clickLangChange.VB_Description = "Callback for langDrop onAction: Change the language of the designer"

    'langId is the language code
    Dim tradsh As Worksheet
    Dim wb As Workbook

    BusyApp

    On Error GoTo ExitLang

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets(TRADTABLESHEET)
    tradsh.Range(RNG_FileLang).Value = langId

    'Update all the labels on the ribbon by reloading it
    UpdateLabels

    'Translate elements in the worksheets
    TranslateWbElmts langId

ExitLang:
    NotBusyApp
End Sub



Private Sub TranslateWbElmts(Byval langId As String)

    Dim wb As Workbook
    Dim pass As IPasswords
    Dim drop As IDropdownLists
    Dim sh As Worksheet
    Dim pass As IPasswords
    Dim hRng As Range
    Dim trads As ITranslation
    Dim selectValue As String

    Set pass = Passwords.Create(wb.Worksheets(PASSSHEETNAME))
    Set trads = Translation.Create(wb.Worksheets(TRADTABLESHEET).ListObjects(1), langId)
    Set drop = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    selectValue = trads.TranslatedValue("selectValue")

    For sh in wb.Worksheets

        'Update elements in the disease worksheet
        If sh.Cells(2, 4).Value = "DISSHEET" Then
            pass.UnProtect sh

            'Change the headers to the corresponding language
            Set hRng = sh.ListObjects(1).HeaderRowRange
            hRng.Cells(1, 1).Value = trads.TranslatedValue("varName")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("varLabel")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("varChoice")
            hRng.Cells(1, 4).Value = trads.TranslatedValue("choiceVal")
            hRng.Cells(1, 5).Value = trads.TranslatedValue("varStatus")
            hRng.Cells(1, 6).Value = trads.TranslatedValue("varVis")

            'Change the dropdown values for the columns status and visibility
            With sh.ListObjects(1)
                
                'variable status
                drop.SetValidation cellRng:=.ListColumns(5).DataBdoyRange, _ 
                                   listName:="__var_status_" & LCase(langId), _
                                   alertType:="error", _ 
                                   message:=selectValue
                
                'variable visibility
                drop.SetValidation cellRng:=.ListColumns(6).DataBodyRange, _ 
                                   listName:="__var_status_" & LCase(langId), _ 
                                   alertType:="error", message:=selectValue
            End With

            pass.Protect sh


        'Update columns in the variable worksheet
        
        ElseIf sh.Name = "Variables" Then


            Set hRng = sh.ListObjects(1).HeaderRowRange
            
            pass.Unprotect sh
            
            hRng.Cells(1, 1).Value = trads.TranslatedValue("varName")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("varLabel")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("defChoice")
            hRng.Cells(1, 4).Value = trads.TranslatedValue("choiceVal")
            hRng.Cells(1, 5).Value = trads.TransatedValue("defStatus")
            hRng.Cells(1, 6).Value = trads.TranslatedValue("comments")

            'Variable status validation
            drop.SetValidation cellRng:= sh.ListObjects(1).ListColumns(5).DataBodyRange, _ 
                               listName:="__var_status_" & LCase(langId), _
                               alertType:="error", message:=selectValue

            pass.Protect sh

        ElseIf sh.Name = "Choices"

            Set hRng = sh.ListObjects(1).HeaderRowRange
            
            pass.UnProtect sh

            hRng.Cells(1, 1).Value = trads.TranslatedValue("varName")
            hRng.Cells(1, 2).Value = trads.TranslatedValue("varLabel")
            hRng.Cells(1, 3).Value = trads.TranslatedValue("defChoice")
            

            pass.Protect sh
        End If

    Next



End Sub