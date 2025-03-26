Attribute VB_Name = "ImportForm"
'@IgnoreModule SheetAccessedUsingString
Option Explicit

'@Folder("Import/Export")

'Sub for functions on the import form
'Write the path to the new setup file to be imported
Public Sub NewSetupPath()
    Dim io As IOSFiles
    Set io = OSFiles.Create()
    'Load a setup file
    io.LoadFile "*.xlsb"
    If io.HasValidFile() Then [Imports].LabPath.Caption = "Path: " & io.File()
End Sub

'Import anaother setup to the new one

'speed app
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


Public Sub ImportOrCleanSetup()
    Dim importDict As Boolean
    Dim importChoi As Boolean
    Dim importExp As Boolean
    Dim importAna As Boolean
    Dim importTrans As Boolean
    Dim importPath As String
    Dim progObj As Object 'Progress label
    Dim importObj As ISetupImport 'Import Object
    Dim sheetsList As BetterArray
    Dim actsh As Worksheet
    Dim infoText As String
    Dim doLabel As String
    Dim pass As IPasswords
    Dim wb As Workbook
    Dim labValue As String
    Dim conformityCheck As Boolean
    Dim counter As Long

    BusyApp
    On Error GoTo errHand
    Application.Cursor = xlWait

    Set actsh = ActiveSheet
    Set wb = ThisWorkbook
    importDict = [Imports].DictionaryCheck.Value
    importChoi = [Imports].ChoiceCheck.Value
    importExp = [Imports].ExportsCheck.Value
    importAna = [Imports].AnalysisCheck.Value
    importTrans = [Imports].TranslationsCheck.Value
    importPath = Application.WorksheetFunction.Trim( _
                Replace([Imports].LabPath, "Path: ", vbNullString))
    conformityCheck = [Imports].ConformityCheck.Value

    Set progObj = [Imports].LabProgress
    Set pass = Passwords.Create(ThisWorkbook.Worksheets("__pass"))
    doLabel = [Imports].DoButton.Caption
    'freeze the pane for modifications
    progObj.Caption = vbNullString
    Set importObj = SetupImport.Create(importPath, progObj)

    'Check import to be sure everything is fine (At least one import has to be made
    'and the file is correct (without missing parts)
    importObj.check importDict, importChoi, importExp, importAna, _
                    importTrans, cleanSetup:=(doLabel = "Clear")

    'Stop import if checks are not valid
    labValue = progObj.Caption
    'Exit the sub in case of error, without proceeding
    If InStr(1, labValue, "Error") > 0 Then Exit Sub
    Set sheetsList = New BetterArray

    'Add the sheets to import if required
    If importDict Then sheetsList.Push "Dictionary"
    If importChoi Then sheetsList.Push "Choices"
    If importExp Then sheetsList.Push "Exports"
    If importAna Then sheetsList.Push "Analysis"
    If importTrans Then sheetsList.Push "Translations"

    Select Case doLabel
    Case "Import"
        importObj.Import pass, sheetsList
         'Check the conformity of current setup file for errors
        If [Imports].ConformityCheck.Value Then CheckTheSetup
        infoText = "Import Done!"
    Case "Clear"
        If (MsgBox("Do you really want to clear the setup?", vbYesNo, "Confirmation") = vbYes) Then
        
            importObj.Clean pass, sheetsList
            
            'Automatically resize tables in the worksheet
            For counter = sheetsList.LowerBound To sheetsList.UpperBound
                BusyApp
                EventsRibbon.ManageRows sheetName:=(sheetsList.Item(counter)), del:=True, allAnalysis:=True
                BusyApp
            Next

            BusyApp
            'Automatically clean the checking worksheet
            On Error Resume Next
            wb.Worksheets("__checkRep").Cells.Clear
            On Error GoTo 0
            infoText = "Setup cleared!"
        Else
            infoText = "Aborted!"
        End If
    End Select

    DoEvents

    'If there is a checking done, no need to add new message
    If Not conformityCheck Then
        MsgBox infoText
        progObj.Caption = infoText
        actsh.Activate
    Else
        [Imports].Hide
        ThisWorkbook.Worksheets("__checkRep").Activate
    End If

    'Fire EnterAnalysis Event to update dropdowns on analysis sheet (force updates)
    EventsAnalysis.EnterAnalysis forceUpdate:=True

    'Call events related to workbook opening
    EventsGlobal.OpenedWorkbook
errHand:
    NotBusyApp
    Application.Cursor = xlDefault
End Sub


