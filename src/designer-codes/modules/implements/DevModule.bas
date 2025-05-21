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

            codeObject.Export outPath
        Else
            Debug.Print moduleName & "not found in current workbook"
        End If

    End Select

End Sub

'Import the path to the class and modules for the building Process

'scope takes two values
'1 for modules
'2 for classes
Private Sub ImportFolder(Optional ByVal scope As Byte = 1)
    Dim io As IOSFiles
    Dim sh As Worksheet
    Dim rng As Range
    Dim rngName As String
    Dim secondFolder As String
    Dim secondRngName As String
    Dim secondRng As Range 'second range for interfaces and test codes

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    rngName = Switch(scope = 1, "RNG_MODULES_CODES_FOLDER", _
                     scope = 2, "RNG_CLASS_CODES_FOLDER", _
                     True, "RNG_MODULES_CODES_FOLDER")
    secondRngName = Switch(scope = 1, "RNG_TEST_MODULES_FOLDER", _
                           scope = 2, "RNG_CLASS_INTERFACE_FOLDER", _
                           True, vbNullString)

    secondFolder = Switch(scope = 1, "tests", _
                     scope = 2, "interfaces", _
                     True, "tests")


    Set rng = sh.Range(rngName)
    Set secondRng = sh.Range(secondRngName)

    Set io = OSFiles.Create()
    io.LoadFolder
    If io.HasValidFolder() Then
        rng.Value = io.Folder() & Application.PathSeparator & "implements"
       secondRng.Value = io.Folder() & Application.PathSeparator & secondFolder
    End If
End Sub

'Helper for SaveCodes
Private Sub SaveOneFolder(ByVal listName As String, outDir As String, scope As Byte, outputAs As Byte)

    Dim codesList As BetterArray
    Dim sh As Worksheet
    Dim codeName As String
    Dim counter As Long

    Set codesList = New BetterArray
    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    codesList.FromExcelRange sh.ListObjects(listName).DataBodyRange()

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = Application.WorksheetFunction.Trim(codesList.Item(counter))
        If codeName <> vbNullString Then TransferCode codeName, outDir, scope:=scope, outputAs:=outputAs
    Next
End Sub

'Import to folders
'outputAs can have 2 values
'1- for Import
'2- for export

'Scope has two values
'1- for modules
'2- for classes
Private Sub SaveCodes(Optional ByVal outputAs As Byte = 1, Optional ByVal scope As Byte = 1)

    Dim sh As Worksheet
    Dim outDir As String
    Dim rngName As String
    Dim listName As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    rngName = Switch(scope = 1, "RNG_MODULES_CODES_FOLDER", _
                     scope = 2, "RNG_CLASS_CODES_FOLDER")
    listName = Switch(scope = 1, "modulesList", scope = 2, "classList")
    'Output directory
    outDir = sh.Range(rngName).Value
    'Be sure the path exists on the current computer before proceeding, if not, exit
    If Dir(outDir & "*", vbDirectory) = vbNullString Then Exit Sub

    SaveOneFolder listName, outDir, scope:=scope, outputAs:=outputAs

    'Second part to Import/Export
    listName = Switch(scope = 1, "testModulesList", scope = 2, "classInterfacesList")
    rngName = Switch(scope = 1, "RNG_TEST_MODULES_FOLDER", _
                     scope = 2, "RNG_CLASS_INTERFACE_FOLDER")
    outDir = sh.Range(rngName).Value

    SaveOneFolder listName, outDir, scope:=scope, outputAs:=outputAs
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

'Add codes to some components
'@EntryPoint
Private Sub CopyCodes(ByVal importModName As String, ByVal exportCodeName As String)
    Dim codeContent As String
    Dim codeMod As Object
    Dim vbProj As Object
    Dim vbComp As Object

    Set vbProj = ThisWorkbook.VBProject

    'Extract the code from the actual vbProject
    With vbProj
        With .VBComponents(importModName).CodeModule
            codeContent = .Lines(1, .CountOfLines)
        End With
    End With

    'Export codeModule
    Set vbComp = vbProj.VBComponents(exportCodeName)
    Set codeMod = vbComp.CodeModule

    'Adding to the export codeModule
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'@Description("Hide some worksheets before deployment")
'@EntryPoint
Public Sub PrepareToDeployment()
Attribute PrepareToDeployment.VB_Description = "Hide some worksheets before deployment"

    'List of sheets to Hide
    Dim sheetsList As BetterArray
    Dim sh As Worksheet
    Dim mainsh As Worksheet
    Dim counter As Long
    Dim wb As Workbook

    Set sheetsList = New BetterArray
    Set wb = ThisWorkbook

    sheetsList.Push "Dictionary", "Choices", "Analysis", "Exports", _
                    "Translations", "__pass", "__formula", "Dev", "__names"

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        Set sh = wb.Worksheets(sheetsList.Item(counter))
        sh.Visible = xlSheetHidden
    Next

    'Add a worksheet change event to the main sheet
    Set mainsh = wb.Worksheets("Main")

    'Add commands on Shapes of the main sheet
    With mainsh
        .Shapes("SHP_LoadDico").OnAction = "LoadFileDic"
        .Shapes("SHP_LoadGeo").OnAction = "LoadGeoFile"
        .Shapes("SHP_LinelistPath").OnAction = "LinelistDir"
        .Shapes("SHP_CtrlNouv").OnAction = "Control"
        .Shapes("SHP_TempFile").OnAction = "LoadTemplateFile"
    End With

    'Add codes to elements on the actual designer
    CopyCodes "EventsMainSheet", mainsh.codeName
    CopyCodes "EventsDesignerWorkbook", wb.codeName
    CopyCodes "FormLogicShowHide", "F_ShowHideLL"
    CopyCodes "FormLogicShowHidePrint", "F_ShowHidePrint"
    CopyCodes "FormLogicGeo", "F_Geo"
    CopyCodes "FormLogicExport", "F_Export"
    CopyCodes "FormLogicAdvanced", "F_Advanced"
    CopyCodes "FormLogicExportMigration", "F_ExportMig"
    CopyCodes "FormLogicImportRep", "F_ImportRep"
    CopyCodes "FormLogicShowVarLabels", "F_ShowVarLabels"

End Sub

'Report Import or export
Private Sub ReportSave(Optional ByVal outputAs As Byte = 1, Optional ByVal scope As Byte = 1)
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim Lo As listObject
    Dim saveName As String
    Dim folderName As String
    Dim phraseToWrite As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    Set Lo = sh.ListObjects("logImports")

    saveName = Switch(outputAs = 1, "Imported ", outputAs = 2, "Exported ", True, "Saved: ")
    folderName = Switch(scope = 1, "Modules using path: " & sh.Range("RNG_MODULES_CODES_FOLDER").Value, _
                        scope = 2, "Classes using path: " & sh.Range("RNG_CLASS_CODES_FOLDER").Value, _
                        True, "<folder>:")

    phraseToWrite = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & saveName & folderName

    Set cellRng = Lo.Range.Cells(1, 1)

    Do While Not IsEmpty(cellRng)
        Set cellRng = cellRng.Offset(1)
    Loop

    cellRng.Value = phraseToWrite
End Sub

'Description("Remove the unnecessary modules/classes")
'@EntryPoint
Public Sub RemoveSub()
    Dim codesList As BetterArray
    Dim sh As Worksheet
    Dim codecounter As Long
    Dim locounter As Long
    Dim componentObject As Object
    Dim codeObject As Object
    Dim wb As Workbook
    Dim LoList As BetterArray
    Dim excludesList As BetterArray 'List of modules to exclude from removing process
    Dim moduleName As String

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(DEVSHEETNAME)
    'Modules list
    Set codesList = New BetterArray
    Set LoList = New BetterArray
    Set excludesList = New BetterArray

    LoList.Push "modulesList", "classList", "testModulesList", "classInterfacesList"
    excludesList.Push "EventsDesignerRibbon", "DevModule", "DropdownLists", "IDropdownLists", _
                      "BetterArray", "OSFiles", "IOSFiles", "LLGeo", "ILLGeo", "TranslationObject", "ITranslationObject", _
                      "DesTranslation", "IDesTranslation", "Main", "IMain", "CustomTable", "ICustomTable", _
                      "LLTranslations", "ILLTranslations", "LLPasswords", "ILLPasswords", "LLdictionary", "ILLdictionary", _
                      "DataSheet", "IDataSheet"

    For locounter = LoList.LowerBound To LoList.UpperBound
        codesList.FromExcelRange sh.ListObjects(LoList.Item(locounter)).DataBodyRange
        For codecounter = codesList.LowerBound To codesList.UpperBound
            On Error Resume Next
                moduleName = codesList.Item(codecounter)
                Set codeObject = wb.VBProject.VBComponents(moduleName)
                Set componentObject = wb.VBProject.VBComponents
                'remove the module from this
                If Not excludesList.Includes(moduleName) Then componentObject.Remove codeObject
                Set codeObject = Nothing
            On Error GoTo 0
        Next
    Next
End Sub
