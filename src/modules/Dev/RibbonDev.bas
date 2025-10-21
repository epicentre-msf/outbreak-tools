Attribute VB_Name = "RibbonDev"
Option Explicit

'@Folder("Dev")
'@ModuleDescription("Ribbon callbacks coordinating development workflows")
'@depends Development, CustomTable, ICustomTable, Passwords, OSFiles, IOSFiles
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing, UseMeaningfulName

Private devManager As IDevelopment

Private Const DEV_SHEET_NAME As String = "Dev"
Private Const CODE_SHEET_NAME As String = "Codes"
Private Const PASS_SHEET_NAME As String = "__pass"
Private Const PROMPT_TITLE As String = "Development"

'@section Ribbon callbacks
'===============================================================================
'@EntryPoint
'@Description("Initialise development environment - logic provided by consuming workbook")
Public Sub clickDevInitialize(ByVal control As IRibbonControl)
    'Reserved for workbook-specific initialisation logic.
End Sub

'@EntryPoint
'@Description("Select root code folder and populate Dev named ranges")
Public Sub clickDevFolder(ByVal control As IRibbonControl)
    On Error GoTo Handler

    Dim devSheet As Worksheet
    Set devSheet = DevSheet()

    Dim io As IOSFiles
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    Dim rootPath As String
    rootPath = io.Folder()

    Dim sep As String
    sep = Application.PathSeparator

    devSheet.Range("ModulesCodes").Value = rootPath & sep & "src" & sep & "modules"
    devSheet.Range("TestsCodes").Value = rootPath & sep & "src" & sep & "tests"
    devSheet.Range("ClassesImplementation").Value = rootPath & sep & "src" & sep & "classes"
    Exit Sub

Handler:
    MsgBox "Unable to configure code folders: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Import modules and classes declared on the Dev tables")
Public Sub clickDevImport(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    On Error GoTo Handler
    manager.ImportAll
    Exit Sub

Handler:
    MsgBox "Import failed: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Export modules and classes declared on the Dev tables")
Public Sub clickDevExport(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    On Error GoTo Handler
    manager.ExportAll
    Exit Sub

Handler:
    MsgBox "Export failed: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Open the VBA editor window")
Public Sub clickDevVBE(ByVal control As IRibbonControl)
    Application.VBE.MainWindow.Visible = True
End Sub

'@EntryPoint
'@Description("Deploy workbook protections and hide Dev artefacts")
Public Sub clickDevDeploy(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    Dim pass As IPasswords
    Set pass = ResolvePasswords()
    If pass Is Nothing Then
        MsgBox "Passwords sheet '" & PASS_SHEET_NAME & "' not found. Cannot deploy.", vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    On Error GoTo Handler
    manager.Deploy pass
    Exit Sub

Handler:
    MsgBox "Deployment failed: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Add default rows to each registered development table")
Public Sub clickDevAddRows(ByVal control As IRibbonControl)
    UpdateTables addRows:=True
End Sub

'@EntryPoint
'@Description("Resize development tables by trimming data rows")
Public Sub clicDevResize(ByVal control As IRibbonControl)
    UpdateTables addRows:=False
End Sub

'@EntryPoint
'@Description("Copy module code into mapped forms")
Public Sub clicDevAddFormTable(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    On Error GoTo Handler
    manager.AddFormsCodes
    Exit Sub

Handler:
    MsgBox "Unable to copy form code: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Create a new forms mapping table")
Public Sub clickDevAddFormTable(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    On Error GoTo Handler
    Dim created As ListObject
    Set created = manager.AddFormsTable
    If created Is Nothing Then Exit Sub
    Exit Sub

Handler:
    MsgBox "Unable to create forms table: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Create a new classes table (general or tests)")
Public Sub clickDevAddClassTable(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    Dim includeTests As Boolean
    includeTests = (MsgBox("Add a tests classes table?", vbYesNo + vbQuestion + vbDefaultButton2, PROMPT_TITLE) = vbYes)

    On Error GoTo Handler
    Dim created As ListObject
    Set created = manager.AddClassTable(includeTests)
    If created Is Nothing Then Exit Sub
    Exit Sub

Handler:
    MsgBox "Unable to create classes table: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@EntryPoint
'@Description("Create a new modules table (general or tests)")
Public Sub clickDevAddModulesTable(ByVal control As IRibbonControl)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Call EnsureCodeSheet(manager)

    Dim includeTests As Boolean
    includeTests = (MsgBox("Add a tests modules table?", vbYesNo + vbQuestion + vbDefaultButton2, PROMPT_TITLE) = vbYes)

    On Error GoTo Handler
    Dim created As ListObject
    Set created = manager.AddModuleTable(includeTests)
    If created Is Nothing Then Exit Sub
    Exit Sub

Handler:
    MsgBox "Unable to create modules table: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub


'@section Helpers
'===============================================================================
Private Function EnsureDevelopment() As IDevelopment
    On Error GoTo Handler

    If devManager Is Nothing Then
        Dim devSheet As Worksheet
        Set devSheet = DevSheet()

        Dim codeSheet As Worksheet
        Set codeSheet = TryWorksheet(CODE_SHEET_NAME)

        If codeSheet Is Nothing Then
            Set devManager = Development.Create(devSheet)
        Else
            'Reuse stored code worksheet when available so tables remain on same sheet.
            Set devManager = Development.Create(devSheet, codeSheet)
        End If
    End If

    Set EnsureDevelopment = devManager
    Exit Function

Handler:
    MsgBox "Unable to initialise Development manager: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
    Set EnsureDevelopment = Nothing
End Function

Private Function DevSheet() As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
        Set sheet = ThisWorkbook.Worksheets(DEV_SHEET_NAME)
    On Error GoTo 0

    If sheet Is Nothing Then
        Err.Raise vbObjectError + 513, PROMPT_TITLE, "Worksheet '" & DEV_SHEET_NAME & "' is required"
    End If

    Set DevSheet = sheet
End Function

Private Function TryWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set TryWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function EnsureCodeSheet(ByVal manager As IDevelopment) As Worksheet
    Dim sheet As Worksheet
    Set sheet = manager.CodeWorksheet

    If sheet Is Nothing Then
        Dim workbookSheet As Worksheet
        Set workbookSheet = TryWorksheet(CODE_SHEET_NAME)
        If Not workbookSheet Is Nothing Then
            On Error Resume Next
                Set sheet = manager.AddCodeSheets(workbookSheet.Name)
            On Error GoTo 0
        End If
    End If

    If sheet Is Nothing Then
        Set sheet = DevSheet()
    End If

    'Return the resolved sheet so callers can operate against it.
    Set EnsureCodeSheet = sheet
End Function

Private Function ResolvePasswords() As IPasswords
    Dim passSheet As Worksheet
    On Error Resume Next
        Set passSheet = ThisWorkbook.Worksheets(PASS_SHEET_NAME)
    On Error GoTo 0

    If passSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolvePasswords = Passwords.Create(passSheet)
    On Error GoTo 0
End Function

Private Sub UpdateTables(ByVal addRows As Boolean)
    Dim manager As IDevelopment
    Set manager = EnsureDevelopment()
    If manager Is Nothing Then Exit Sub

    Dim targetSheet As Worksheet
    Set targetSheet = EnsureCodeSheet(manager)

    Dim pass As IPasswords
    Set pass = ResolvePasswords()

    On Error GoTo Cleanup
    If Not pass Is Nothing Then pass.UnProtect targetSheet

    Dim lo As ListObject
    For Each lo In targetSheet.ListObjects
        Dim table As ICustomTable
        Set table = CustomTable.Create(lo)
        If addRows Then
            'Pad tables with one extra row to speed up data entry.
            table.AddRows
        Else
            'Reset tables back to structural rows only.
            table.RemoveRows totalCount:=0
        End If
    Next lo

Cleanup:
    If Not pass Is Nothing Then pass.Protect targetSheet
    If Err.Number <> 0 Then
        MsgBox "Unable to update tables: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub
