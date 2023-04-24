Attribute VB_Name = "DevModule"

Option Explicit

'@Folder("Development Procedures")

'Working with import/export of the code in the designer
'Scope can take 2 values :
'1- For modules
'2- For classes

'outputAs can take 2 values:
'1 for import
'2 for export

'outDir is the output directory
'moduleName is the name of the module in the output directory
Private Const DEVSHEETNAME As String = "Dev"

Private Sub TransferCode(ByVal moduleName As String, ByVal outDir As String, Optional ByVal scope As Byte = 1, Optional ByVal outputAs As Byte = 1)

    Dim codeObject As Object
    Dim componentObject As Object
    Dim Wkb As Workbook
    Dim outPath As String
    Dim sep As String


    Set Wkb = ThisWorkbook                       'Output workbook
    sep = Application.PathSeparator

    'get all the values within the actual module to transfer
    Select Case scope
    Case 1
        outPath = outDir & sep & moduleName & ".bas"
    Case 2
        outPath = outDir & sep & moduleName & ".cls"
    End Select

    'I need to import/export classes to keep their attribute. (self instanciation, etc.)

    Select Case outputAs

        'Import the module in the current file
    Case 1
        On Error Resume Next
        Set codeObject = Wkb.VBProject.VBComponents(moduleName)
        Set componentObject = Wkb.VBProject.VBComponents

        'remove the module from this
        componentObject.Remove codeObject
        On Error GoTo 0

        'Be sure here that the path and files exist before import
        componentObject.Import outPath

    Case 2
        'Export the module to the output directory
        On Error Resume Next
        Set codeObject = Wkb.VBProject.VBComponents(moduleName)
        On Error GoTo 0

        If Not (codeObject Is Nothing) Then

            'Remove the file if its exists
            On Error Resume Next
            Kill outPath
            On Error GoTo 0

            codeObject.export outPath
        Else
            Debug.Print moduleName & "not found in current workbook"
        End If

    End Select

End Sub

'Import the path to the class and modules for the building Process

'scope takes two values
'1 for modules
'2 for classes
Public Sub ImportFolder(Optional ByVal scope As Byte = 1)
    Dim io As IOSFiles
    Dim sh As Worksheet
    Dim rng As Range
    Dim rngName As String
    Dim secondFolder As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    rngName = Switch(scope = 1, "RNG_MODULES_CODES_FOLDER", _
                     scope = 2, "RNG_CLASS_CODES_FOLDER", _
                     True, "RNG_MODULES_CODES_FOLDER")
    
    secondFolder  = Switch(scope = 1, "tests", _
                     scope = 2, "interfaces", _
                     True, "tests")


    Set rng = sh.Range(rngName)
    Set io = OSFiles.Create()
    io.LoadFolder
    If io.HasValidFolder() Then 
        rng.Value = io.Folder() & Application.PathSeparator & "codes"

        If scope = 1 Then
            rng.Offset(2).Value = io.Folder() & Application.PathSeparator & "tests"
        Else
            rng.Offset(2).Value = io.Folder() & Application.PathSeparator & "interfaces"
        End If
End Sub

'Import to folders
'outputAs can have 2 values
'1- for Import
'2- for export

'Scope has two values
'1- for modules
'2- for classes
Public Sub SaveCodes(Optional ByVal outputAs As Byte = 1, Optional ByVal scope As Byte = 1)
    Dim codesList As BetterArray
    Dim sh As Worksheet
    Dim counter As Long
    Dim outDir As String
    Dim codeName As String
    Dim rngName As String
    Dim listName As String


    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    rngName = Switch(scope = 1, "RNG_MODULES_FOLDER", _
                     scope = 2, "RNG_CLASS_FOLDER", _
                     True, "RNG_MODULES_FOLDER")

    listName = Switch(scope = 1, "modulesList", _
                      scope = 2, "classList", _
                      True, "modulesList")

    Set codesList = New BetterArray
    outDir = sh.Range(rngName).Value

    'Be sure the path exists on the current computer before proceeding, if not, exit
    If Dir(outDir & "*", vbDirectory) = vbNullString Then Exit Sub

    codesList.FromExcelRange sh.ListObjects(listName).DataBodyRange()

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = Application.WorksheetFunction.Trim(codesList.Item(counter))
        If codeName <> vbNullString Then TransferCode codeName, outDir, scope:=scope, outputAs:=outputAs
    Next

End Sub

'@Description("Import Codes into the designer")
'@EntryPoint
Public Sub ImportCodes()
Attribute ImportCodes.VB_Description = "Import Codes into the designer"
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    If Not sh.ProtectContents Then
        If MsgBox("Are you sure you want to import the codes ?", vbYesNo) = vbYes Then

            'Import modules
            SaveCodes outputAs:=1, scope:=1
            ReportSave outputAs:=1, scope:=1

            'Import classes
            SaveCodes outputAs:=1, scope:=2
            ReportSave outputAs:=1, scope:=2

            MsgBox "Done!"
            sh.Range("RNG_INFO").Value = "Finished Imports"
        End If
    Else
        sh.Range("RNG_INFO").Value = "Unlock the worksheet before proceeding"
    End If

End Sub

'@Description("Export codes into the designer")
'@EntryPoint
Public Sub ExportCodes()
Attribute ExportCodes.VB_Description = "Export codes into the designer"
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    If Not sh.ProtectContents Then
        If MsgBox("Are you sure to export the codes?", vbYesNo) = vbYes Then
            'Export modules
            SaveCodes outputAs:=2, scope:=1
            ReportSave outputAs:=2, scope:=1

            'Export classes
            SaveCodes outputAs:=2, scope:=2
            ReportSave outputAs:=2, scope:=2

            MsgBox "Done!"
            sh.Range("RNG_INFO").Value = "Finished Exports"
        End If
    Else
        sh.Range("RNG_INFO").Value = "Unlock the worksheet before proceeding"
    End If

End Sub

'@Description("Import class folder path")
'@EntryPoint
Public Sub ImportClassFolder()
Attribute ImportClassFolder.VB_Description = "Import class folder path"
    ImportFolder scope:=2
End Sub

'@Description("Import module folder path")
'@EntryPoint
Public Sub ImportModuleFolder()
Attribute ImportModuleFolder.VB_Description = "Import module folder path"
    ImportFolder scope:=1
End Sub

'@Description("Hide some worksheets before deployment")
'@EntryPoint
Public Sub PrepareToDeployment()
Attribute PrepareToDeployment.VB_Description = "Hide some worksheets before deployment"

    Const WBMODULENAME As String = "EventsDesignerWorkbook"
    Const MAINMODULENAME As String = "EventsMainSheet"

    'List of sheets to Hide
    Dim sheetsList As BetterArray
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim mainsh As Worksheet
    Dim counter As Long
    'Elements to add Envent code to a worksheet
    Dim mainCodeContent As String
    Dim wbCodeContent As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim mainCodeMod As Object
    Dim wbCodeMod As Object

    Set sheetsList = New BetterArray
    Set wb = ThisWorkbook

    sheetsList.Push "Dictionary", "Choices", "Analysis", "Exports", _
                    "Translations", "__pass", "__formula"

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        Set sh = wb.Worksheets(sheetsList.Item(counter))
        sh.Visible = xlSheetHidden
    Next

    'Add a worksheet change event to the main sheet
    Set mainsh = wb.Worksheets("Main")

    With wb.VBProject
        With .VBComponents(MAINMODULENAME).CodeModule
            mainCodeContent = .Lines(1, .CountOfLines)
        End With
        With .VBComponents(WBMODULENAME).CodeModule
            wbCodeContent = .Lines(1, .CountOfLines)
        End With
    End With

    With wb
        'main code module
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(mainsh.codeName)
        Set mainCodeMod = vbComp.CodeModule

        'wb code module
        Set vbComp = vbProj.VBComponents(.codeName)
        Set wbCodeMod = vbComp.CodeModule
    End With

    'Adding the code to main sheet
    With mainCodeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString mainCodeContent
    End With

    'Adding the code to the workbook
    With wbCodeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString wbCodeContent
    End With
End Sub

'Report Import or export
Private Sub ReportSave(Optional ByVal outputAs As Byte = 1, Optional ByVal scope As Byte = 1)
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim Lo As ListObject
    Dim saveName As String
    Dim folderName As String
    Dim phraseToWrite As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    Set Lo = sh.ListObjects("logImports")

    saveName = Switch(outputAs = 1, "Imported ", outputAs = 2, "Exported ", True, "Saved: ")
    folderName = Switch(scope = 1, "Modules using path: " & sh.Range("RNG_MODULES_FOLDER").Value, _
                        scope = 2, "Classes using path: " & sh.Range("RNG_CLASS_FOLDER").Value, _
                        True, "<folder>:")

    phraseToWrite = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & saveName & folderName

    Set cellRng = Lo.Range.Cells(1, 1)

    Do While Not IsEmpty(cellRng)
        Set cellRng = cellRng.Offset(1)
    Loop

    cellRng.Value = phraseToWrite
End Sub
