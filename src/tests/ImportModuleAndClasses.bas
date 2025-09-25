Attribute VB_Name = "ImportModuleAndClasses"
Option Explicit

'@Folder("Dev")
'@IgnoreModule

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

Private Const MODULECODESRANGE As String = "ModulesCodes"
Private Const CLASSCODESRANGE As String = "ClassesImplementation"
Private Const TESTCODESRANGE As String = "TestsCodes"

Private ClassFolders As String
Private TestFolders As String
Private ModuleFolders As String


Private Enum ImportedFileScope
    moduleImport = 1
    classImport = 2
End Enum

Private Enum TransfertFileScope
    ImportIntoFile = 1
    ExportToPath = 2
End Enum

'Transfert code from a directory to this file (or from this file to a directory)
Private Sub TransferCode(ByVal moduleName As String, ByVal outDir As String, _
                        Optional ByVal scope As Byte = ImportedFileScope.moduleImport, _
                        Optional ByVal outputAs As Byte = TransfertFileScope.ImportIntoFile)

    Dim codeObject As Object
    Dim componentObject As Object
    Dim Wkb As Workbook
    Dim outPath As String
    Dim sep As String

    Set Wkb = ThisWorkbook                       'Output workbook
    sep = Application.PathSeparator

    'get all the values within the actual module to transfer
    Select Case scope
    Case moduleImport
        outPath = outDir & sep & moduleName & ".bas"
    Case classImport
        outPath = outDir & sep & moduleName & ".cls"
    End Select

    'I need to import/export classes to keep their attribute. (self instanciation, etc.)

    Select Case outputAs

        'Import the module in the current file
    Case ImportIntoFile
        On Error Resume Next
        Set codeObject = Wkb.VBProject.VBComponents(moduleName)
        Set componentObject = Wkb.VBProject.VBComponents

        'remove the module from this
        componentObject.Remove codeObject
        On Error GoTo 0

        'Be sure here that the path and files exist before import
        componentObject.Import outPath

    Case ExportToPath
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

'@Description("Callback for btnResize onAction")
'@EntryPoint
Public Sub clickRibbonFolder(ByRef Control As IRibbonControl)
    Dim io As IOSFiles
    Dim sh As Worksheet
    Dim sep As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    sep = Application.PathSeparator

    Set io = OSFiles.Create()
    io.LoadFolder
    If io.HasValidFolder() Then
       sh.Range(MODULECODESRANGE).Value = io.Folder() & sep & "src" & sep & "modules"
       sh.Range(TESTCODESRANGE).Value = io.Folder() & sep & "src" & sep & "tests"
       sh.Range(CLASSCODESRANGE).Value = io.Folder() & sep & "src" & sep & "classes"
    End If
End Sub

Private Sub ResolveOutputDirs()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    ClassFolders = sh.Range(CLASSCODESRANGE).Value
    TestFolders = sh.Range(TESTCODESRANGE).Value
    ModuleFolders = sh.Range(MODULECODESRANGE).Value

End Sub

'Helper for SaveCodes
Private Sub SaveOneFolder(ByVal listName As String, outDir As String, scope As Byte, outputAs As Byte, Optional ByVal interFace As Boolean = True)

    Dim sh As Worksheet
    Dim codeName As String
    Dim codeNameInterFace As String
    Dim counter As Long
    Dim colRng As Range


    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    On Error Resume Next
    Set codeRng = sh.ListObjects(listName).ListColumns(1).DataBodyRange
    On Error GoTo 0

    If codeRng Is Nothing Then Exit Sub

    For counter = 1 To colRng.Rows.Count
        codeName = Application.WorksheetFunction.Trim(colRng.Cells(counter, 1).Value)

        If codeName <> vbNullString Then

            TransferCode codeName, outDir, scope:=scope, outputAs:=outputAs

            If interFace And (Cstr(colRng.Cells(counter, 2).Value) = "yes") Then 
                codeNameInterFace = "I" & codeName
                TransferCode codeNameInterFace, outDir, scope:=scope, outputAs:=outputAs
            End If
        End If
    Next
End Sub

'@Description("Callback for btnVBE onAction")
'@EntryPoint
Public Sub clickRibbonVBE(ByRef Control As IRibbonControl)
    Application.VBE.MainWindow.Visible = True
End Sub

'Import to folders
'outputAs can have 2 values
'1- for Import
'2- for export

'Scope has two values
'1- for modules
'2- for classes
Private Sub SaveCodes(Lo As ListObject, Optional ByVal outputAs As Byte = ImportIntoFile)

    Dim sh As Worksheet
    Dim outDir As String
    Dim codeScope As String
    Dim codeFolder As String
    Dim listName As String
    Dim hasInterface As Boolean
    Dim importScope As Byte

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)
    
    ResolveOutputDirs
    
    codeScope = Cstr(Lo.Range.Cells(0, 1).Value)
    codeFolder = Cstr(Lo.Range.Cells(-1, 1).Value)
    listName = Lo.Name

    Select Case codeScope

    Case "tests modules"

        outDir = TestFolders & Application.PathSeparator & "modules"
        importScope = moduleImport

    Case "general modules"

        outDir = ModuleFolders
        importScope = moduleImport

    Case "tests classes"

        outDir = TestFolders & Application.PathSeparator & "classes"
        importScope = classImport

    Case "general classes"
        outDir = ClassFolders 
        importScope = classImport
        hasInterFace = True

    End Select

    outDir = outDir & Application.PathSeparator & codeFolder

    
    If Dir(outDir & "*", vbDirectory) <> vbNullString Then
        SaveOneFolder listName, outDir, importScope, outputAs:=outputAs, interFace:=hasInterface
    End If
   
    ReportSave path:=outDir, outputAs:=outputAs, scope:=importScope
    
End Sub

'@Description("Import Codes into the designer")
'@EntryPoint
Public Sub clickRibbonImport(ByRef Control As IRibbonControl)
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    If Not sh.ProtectContents Then
        If MsgBox("Are you sure you want to import the codes ?", vbYesNo) = vbYes Then

            For Each Lo In sh.ListObjects
                SaveCodes Lo
            Next

            sh.Range("Informations").Value = "Finished Imports At: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    Else
        sh.Range("Informations").Value = "Unlock the worksheet before proceeding"
    End If
End Sub

'@Description("Export codes into the designer")
'@EntryPoint
Public Sub clickRibbonExport(ByRef Control As IRibbonControl)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)

    If Not sh.ProtectContents Then
        If MsgBox("Are you sure to export the codes?", vbYesNo) = vbYes Then
            SaveCodes outputAs:=ExportToPath
            MsgBox "Done!"
            sh.Range("Informations").Value = "Finished Exports"
        End If
    Else
        sh.Range("Informations").Value = "Unlock the worksheet before proceeding"
    End If
End Sub

'Report Import or export
Private Sub ReportSave(Optional ByVal path As String = vbNullString, Optional ByVal outputAs As Byte = 1, Optional ByVal scope As Byte = 1)
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim saveName As String
    Dim folderName As String
    Dim phraseToWrite As String

    Set sh = ThisWorkbook.Worksheets(DEVSHEETNAME)


    saveName = Switch(outputAs = 1, "Imported ", outputAs = 2, "Exported ", True, "Saved: ")
    folderName = Switch(scope = 1, "Modules using path: " & path, scope = 2, "Classes using path: " & path, True, "<folder>:")
    phraseToWrite = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & saveName & folderName

    Set cellRng = sh.Range("Informations").Offset(9)
    Do While Not IsEmpty(cellRng)
        Set cellRng = cellRng.Offset(1)
    Loop

    cellRng.Value = phraseToWrite
End Sub


