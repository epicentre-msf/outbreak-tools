Attribute VB_Name = "MasterSetupExports"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Disease exports and the migration round trip for the master setup workbook.")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'Three doors to the outside world, all built on the msetup classes:
'- ExportToSetup writes one disease as a dictionary workbook a setup imports.
'- ExportForMigration writes every disease of this file into one flat
'  workbook, beside the Translations, Choices and Variables sheets.
'- ImportFlatFile reads such a flat workbook back and recreates or merges
'  every disease worksheet it carries.

Private Const DISEASES_SHEET As String = "Diseases"
Private Const IMPORT_STAGING_SHEET As String = "__dis_import"
Private Const LANGUAGES_DROPDOWN As String = "__data_languages"
Private Const DISEASE_TABLE_WIDTH As Long = 7
Private Const BLOCK_HEADER_ROW As Long = 1
Private Const BLOCK_META_ROW As Long = 2
Private Const BLOCK_TABLE_ROW As Long = 3

'@section Exports
'===============================================================================

'@sub-title Export the active disease worksheet as a setup dictionary workbook.
Public Sub ExportToSetup()
    Dim targetSheet As Worksheet
    Dim io As OSFiles
    Dim exporter As DiseaseExporter
    Dim translationTable As ListObject
    Dim filePath As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox "Select a disease worksheet before exporting it.", vbExclamation + vbOKOnly, "Export"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    Set translationTable = ResolveTranslationsTable()

    Set exporter = BuildExporter()
    filePath = exporter.ExportDisease(io.Folder(), targetSheet, translationTable, _
                                      MasterSetupHelpers.ResolveRibbonTranslations(), _
                                      targetSheet.Name, _
                                      MasterSetupHelpers.SafeValue(targetSheet.Cells(2, 2).Value), _
                                      MasterSetupHelpers.SafeValue(targetSheet.Cells(2, 3).Value))

    MsgBox "Done! The disease file is saved at:" & vbNewLine & filePath, vbInformation + vbOKOnly, "Export"
End Sub

'@sub-title Export every disease of this workbook into one flat migration file.
Public Sub ExportForMigration()
    Dim io As OSFiles
    Dim exporter As DiseaseExporter
    Dim diseaseNames As BetterArray
    Dim filePath As String

    Set diseaseNames = CollectDiseaseNames()
    If diseaseNames.Length = 0 Then
        MsgBox "This workbook carries no disease worksheet to export.", vbExclamation + vbOKOnly, "Export"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    Set exporter = BuildExporter()
    filePath = exporter.ExportForMigration(io.Folder(), ThisWorkbook, diseaseNames)

    MsgBox "Done! The migration file is saved at:" & vbNewLine & filePath, vbInformation + vbOKOnly, "Export"
End Sub

'@section Import
'===============================================================================

'@sub-title Read a flat migration file and recreate or merge its diseases.
Public Sub ImportFlatFile()
    Dim io As OSFiles
    Dim sourceBook As Workbook
    Dim diseasesSheet As Worksheet
    Dim importedCount As Long

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx;*.xlsb;*.xlsm"
    If Not io.HasValidFile() Then Exit Sub

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo CloseSource

    Set diseasesSheet = FindWorksheet(sourceBook, DISEASES_SHEET)
    If diseasesSheet Is Nothing Then
        MsgBox "The selected file carries no '" & DISEASES_SHEET & "' worksheet.", _
               vbExclamation + vbOKOnly, "Import"
        GoTo CloseSource
    End If

    importedCount = ImportMigrationDiseases(sourceBook, ThisWorkbook)

    MsgBox "Done! " & importedCount & " disease worksheet(s) imported.", _
           vbInformation + vbOKOnly, "Import"

CloseSource:
    'Shielded close: the source workbook must not stay open on any path.
    On Error Resume Next
    sourceBook.Close saveChanges:=False
    On Error GoTo 0
End Sub

'@section Import helpers
'===============================================================================

'@sub-title Walk the Diseases sheet blocks of a flat file and import each one.
'@details Public and fully parameterised so a suite can drive the round trip
'without a file picker. ImportFlatFile wraps it with the pickers.
'@param sourceBook Workbook carrying the flat Diseases sheet.
'@param targetBook Workbook receiving the disease worksheets.
'@return Long number of disease blocks imported.
Public Function ImportMigrationDiseases(ByVal sourceBook As Workbook, ByVal targetBook As Workbook) As Long
    Dim diseasesSheet As Worksheet
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim diseaseName As String
    Dim languageTag As String

    Set diseasesSheet = FindWorksheet(sourceBook, DISEASES_SHEET)
    If diseasesSheet Is Nothing Then Exit Function

    lastColumn = diseasesSheet.UsedRange.Columns(diseasesSheet.UsedRange.Columns.Count).Column

    For columnIndex = 1 To lastColumn
        If MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_HEADER_ROW, columnIndex).Value) = "Disease" Then
            diseaseName = MasterSetupHelpers.CleanMasterSheetName( _
                          MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex).Value))
            languageTag = MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex + 1).Value)

            If LenB(diseaseName) > 0 Then
                ImportOneDisease diseasesSheet, columnIndex, diseaseName, languageTag, targetBook
                ImportMigrationDiseases = ImportMigrationDiseases + 1
            End If
        End If
    Next columnIndex
End Function

'@sub-title Import one column block into a fresh or existing disease sheet.
Private Sub ImportOneDisease(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long, _
                             ByVal diseaseName As String, ByVal languageTag As String, _
                             ByVal targetBook As Workbook)

    Dim targetSheet As Worksheet
    Dim stagingTable As ListObject
    Dim manager As DiseaseWorksheetManager
    Dim freshSheet As Boolean

    Set targetSheet = FindWorksheet(targetBook, diseaseName)
    freshSheet = targetSheet Is Nothing

    If freshSheet Then
        Set targetSheet = BuildDiseaseSheet(diseaseName, languageTag, targetBook)
    ElseIf Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        'A sheet of that name that is no disease sheet is left alone.
        Exit Sub
    End If

    If targetSheet.ListObjects.Count = 0 Then Exit Sub

    Set stagingTable = BuildStagingTable(diseasesSheet, startColumn, targetBook)
    If stagingTable Is Nothing Then Exit Sub

    'A fresh sheet takes the block whole; an existing one merges it, with the
    'imported values winning.
    DiseaseImporter.Create().MergeDisease targetSheet.ListObjects(1), stagingTable, _
                                          mergeValues:=Not freshSheet, _
                                          priority:=DiseaseImportPriority_Foreign

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet targetBook, IMPORT_STAGING_SHEET
End Sub

'@sub-title Build a disease sheet for an imported block.
Private Function BuildDiseaseSheet(ByVal diseaseName As String, ByVal languageTag As String, _
                                   ByVal targetBook As Workbook) As Worksheet
    Dim builder As DiseaseSheet
    Dim dropdowns As DropdownLists
    Dim languages As BetterArray

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns( _
                    MasterSetupHelpers.ResolveMasterDropdownsSheet(targetBook))

    Set builder = DiseaseSheet.Create(targetBook, dropdowns, _
                                      MasterSetupHelpers.ResolveRibbonTranslations(targetBook), _
                                      MasterSetupHelpers.ResolveMasterSetupVariables( _
                                      MasterSetupHelpers.ResolveMasterVariablesSheet(targetBook)))

    'A language the target file does not carry falls back to the default.
    Set languages = dropdowns.Values(LANGUAGES_DROPDOWN)
    If Not languages Is Nothing Then
        If Not languages.Includes(languageTag) Then languageTag = vbNullString
    End If

    Set BuildDiseaseSheet = builder.Build(diseaseName, languageTag)
End Function

'@sub-title Copy one block into a staging sheet and answer its ListObject.
Private Function BuildStagingTable(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long, _
                                   ByVal targetBook As Workbook) As ListObject
    Dim stagingSheet As Worksheet
    Dim manager As DiseaseWorksheetManager
    Dim rowCount As Long
    Dim blockRange As Range
    Dim tableRange As Range

    rowCount = CountBlockRows(diseasesSheet, startColumn)
    If rowCount = 0 Then Exit Function

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet targetBook, IMPORT_STAGING_SHEET

    Set stagingSheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
    stagingSheet.Name = IMPORT_STAGING_SHEET
    stagingSheet.Visible = xlSheetHidden

    'Header plus data, copied as values in one assignment.
    Set blockRange = diseasesSheet.Cells(BLOCK_TABLE_ROW, startColumn).Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    Set tableRange = stagingSheet.Range("A1").Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    tableRange.Value = blockRange.Value

    Set BuildStagingTable = stagingSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

'@sub-title Count the data rows of a block, reading the name column.
Private Function CountBlockRows(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long) As Long
    Dim rowIndex As Long
    Dim nameColumn As Long

    'The variable name is the second column of a disease table.
    nameColumn = startColumn + 1
    rowIndex = BLOCK_TABLE_ROW + 1

    Do While LenB(MasterSetupHelpers.SafeValue(diseasesSheet.Cells(rowIndex, nameColumn).Value)) > 0
        CountBlockRows = CountBlockRows + 1
        rowIndex = rowIndex + 1
    Loop
End Function

'@section Shared helpers
'===============================================================================

'@sub-title Build a ready exporter over a fresh workbook manager and guard.
Private Function BuildExporter() As DiseaseExporter
    Set BuildExporter = DiseaseExporter.Create(DiseaseExportWorkbook.Create(), _
                                               ApplicationState.Create(Application))
End Function

'@sub-title Collect the names of every disease worksheet of this workbook.
Private Function CollectDiseaseNames() As BetterArray
    Dim sh As Worksheet

    Set CollectDiseaseNames = New BetterArray
    CollectDiseaseNames.LowerBound = 1

    For Each sh In ThisWorkbook.Worksheets
        If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
            CollectDiseaseNames.Push sh.Name
        End If
    Next sh
End Function

Private Function ResolveTranslationsTable() As ListObject
    Dim translationsSheet As Worksheet

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then Exit Function
    If translationsSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveTranslationsTable = translationsSheet.ListObjects(1)
End Function

Private Function FindWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In targetBook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet
End Function
