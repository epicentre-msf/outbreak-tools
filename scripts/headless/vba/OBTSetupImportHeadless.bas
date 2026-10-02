Attribute VB_Name = "OBTSetupImportHeadless"
Attribute VB_Description = "Import one setup workbook into another with no dialog and no form"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Headless")
'@ModuleDescription("Import one setup workbook into another with no dialog and no form")

Option Explicit

'@description
'ONE HEADLESS SETUP IMPORT.
'
'This module is INJECTED into the target setup workbook and run there. It is
'not part of src/ and no suite imports it: a caller reads it off disk and
'hands it to VBComponents.Import, which is what
'HeadlessBuild.ImportSetupFromWorkbook does.
'
'WHY IT HAS TO RUN INSIDE THE TARGET
'-------------------------------------------------------------------------------
'SetupImport resolves its destination as ThisWorkbook, in seven places. Run
'from a driver workbook it would import the source INTO THE DRIVER. So the
'code that calls it has to live in the workbook being filled, which is why
'this module is injected rather than imported once.
'
'WHAT IT RUNS, AND WHY IN THAT ORDER
'-------------------------------------------------------------------------------
'The sequence is the production one, read off SetupHelpers.ImportSetupData and
'its ExecuteImportOperation, with the form and the dialogs taken out:
'
'  BuildSelectedSheets  the five sheet names the import moves
'  ResolveSetupPasswords  the Passwords over the __pass sheet
'  EnterBusyState  events, screen and calculation, restored at the end
'  SetupImport.Create  the service, over the SOURCE path
'  service.Check  what the source carries against what was asked for
'  service.Import  the move itself
'  PostImportMaintenance  what the production path runs after every import
'
'The conformity check (CheckTheSetup) is deliberately NOT run: it writes a
'report sheet the caller has not asked for, and a headless caller reads the
'return value instead.
'
'NOTHING RAISES OUT OF HERE
'-------------------------------------------------------------------------------
'A raise crossing back into a driver mid-import leaves the target workbook
'busy: events off, calculation manual, screen updating off. Every exit goes
'through the cleanup label, and the outcome is the RETURN VALUE - "OK" or
'"ERROR <number>: <description>". A caller reads that string; it never has to
'read a dialog.

Private Const OUTCOME_OK As String = "OK"


'@Description("Import every data sheet of one setup workbook into this one.")
'@details
'@param sourcePath String. The setup workbook to read, as a full path.
'@return String. OUTCOME_OK, or "ERROR <number>: <description>".
Public Function OBTHeadlessImportSetup(ByVal sourcePath As String) As String
    Dim service As SetupImport
    Dim pass As Passwords
    Dim sheetsList As BetterArray
    Dim busyEntered As Boolean

    On Error GoTo Failed

    If LenB(Trim$(sourcePath)) = 0 Then
        OBTHeadlessImportSetup = "ERROR 0: no source setup path was given"
        Exit Function
    End If

    If Len(Dir$(sourcePath)) = 0 Then
        OBTHeadlessImportSetup = "ERROR 0: the source setup does not exist - " & sourcePath
        Exit Function
    End If

    Set sheetsList = SetupHelpers.BuildSelectedSheets(True, True, True, True, True)
    Set pass = SetupHelpers.ResolveSetupPasswords()

    EventsManager.EnterBusyState calculateOnSave:=False
    busyEntered = True

    Set service = SetupImport.Create(sourcePath)
    service.DisplayPrompts = False
    service.Check True, True, True, True, True, cleanSetup:=False
    service.Import pass, sheetsList

    'What the production path runs after an import lands, so a headless fill
    'leaves the workbook in the same state a hand import does.
    On Error Resume Next
        SetupHelpers.PostImportMaintenance
    On Error GoTo Failed

    OBTHeadlessImportSetup = OUTCOME_OK
    GoTo Cleanup

Failed:
    OBTHeadlessImportSetup = "ERROR " & CStr(Err.Number) & ": " & Err.Description

Cleanup:
    If busyEntered Then
        On Error Resume Next
            EventsManager.ExitBusyState
        On Error GoTo 0
    End If

    Set service = Nothing
    Set pass = Nothing
    Set sheetsList = Nothing
End Function

'@Description("Answer the sheet names this workbook carries, for a caller's check.")
'@return String. The worksheet names, pipe separated.
Public Function OBTHeadlessSheetNames() As String
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
        OBTHeadlessSheetNames = OBTHeadlessSheetNames & sh.Name & "|"
    Next
End Function

'@Description("Answer how many rows one table of this workbook holds.")
'@details
'What a caller checks a fill against: an empty setup answers 0 and a filled
'one answers the row count the source carried.
'@param sheetName String. The worksheet holding the table.
'@param tableName String. The ListObject to measure.
'@return Long. The row count, or -1 when either name does not resolve.
Public Function OBTHeadlessTableRows(ByVal sheetName As String, _
                                     ByVal tableName As String) As Long
    Dim sh As Worksheet
    Dim Lo As ListObject

    OBTHeadlessTableRows = -1

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    On Error Resume Next
        Set Lo = sh.ListObjects(tableName)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Function

    If Lo.DataBodyRange Is Nothing Then
        OBTHeadlessTableRows = 0
    Else
        OBTHeadlessTableRows = Lo.DataBodyRange.Rows.Count
    End If
End Function
