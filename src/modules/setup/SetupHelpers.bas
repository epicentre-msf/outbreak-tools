Attribute VB_Name = "SetupHelpers"

Option Explicit

Private Const PASSSHEETNAME As String = "__pass"
Private Const TRADSHEETNAME As String = "Translations"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const DICTSHEETNAME As String = "Dictionary"


'Start Rows and columns for dictionary, choices, and exports.
Private Const START_ROW_DICTIONARY As Long = 5
Private Const START_ROW_CHOICES As Long = 4
Private Const START_ROW_EXPORTS As Long = 4
Private Const START_COLUMN_DICTIONARY As Long = 1
Private Const START_COLUMN_CHOICES As Long = 1
Private Const START_COLUMN_EXPORTS As Long = 1

'Implement the password protection for the workbook entirely



'@sub-title Add or remove rows to a table
Public Sub ManageRows(ByVal sheetName As String, _
                      Optional ByVal del As Boolean = False)
    Dim part As Object
    Dim targetSheet As Worksheet
    Dim dictSheet As Worksheet
    Dim dict As ILLdictionary
    Dim app As IApplicationState

    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If (targetSheet Is Nothing) Then Exit Sub

    On Error GoTo Handler

    '5 is the start line of the dictionary
    '4 is the start column of the dictionary
    Select Case LCase$(Trim$(sheetName))
    Case "dictionary"
        Set part = LLdictionary.Create(targetSheet START_ROW_DICTIONARY, START_COLUMN_DICTIONARY)
    Case "choices"
        Set part = LLChoices.Create(targetSheet START_ROW_CHOICES , START_COLUMN_CHOICES)
    Case "analysis"
        Set part = Analysis.Create(targetSheet)
    Case "exports"
        Set dictSheet = ThisWorkbook.Worksheets(DICTSHEETNAME)
        Set part = LLExport.Create(targetSheet START_ROW_EXPORTS, START_COLUMN_EXPORTS)
        Set dict = LLdictionary.Create(dictSheet, START_ROW_DICTIONARY, START_COLUMN_DICTIONARY)
    Case Else
        Exit Sub
    End Select

    If Not (part Is Nothing) Then
        app.ApplyBusyState(suppressEvents:=True, calculateOnSave:=False)
        EnsureRowManagement sheetName, del, part, dict
        app.Restore()
    End If

    Exit Sub
Handler:
    On Error Resume Next
    If Not app Is Nothing Then app.Restore()
    If Err.Number <> 0 Then Debug.Print "Manage rows exited with an error: "; Err.Description; Err.Number 
End Sub


'@sub-title Ensure Row Management
Private Sub EnsureRowManagement(ByVal sheetName As String, ByVal del As Boolean, _ 
                                ByVal part As Object, Optional ByVal dict As ILLdictionary)
    
    UnProtectSetupSheet sheetName
    If dict Is Nothing Then
        part.ManageRows del
    Else
        part.ManageRows del, dict
    End If
    ProtectSetupSheet sheetName
End Sub

'@sub-title Unprotect a worksheet
Public Sub UnProtectSetupSheet(ByVal sheetName As String)
    Dim pass As IPasswords
    Set pass = Passwords.Create(ThisWorkbook.Worksheets(PASSSHEETNAME))
    pass.UnProtect sheetName
End Sub

'@sub-title Protect a worksheet
Public Sub ProtectSetupSheet(ByVal sheetName As String)
    Dim pass As IPasswords
    Dim delRow As Boolean

    delRow = Not ((sheetName = TRADSHEETNAME) Or (sheetName = ANALYSISSHEETNAME))

    Set pass = Passwords.Create(ThisWorkbook.Worksheets(PASSSHEETNAME))
    pass.Protect sheetName, allowDeleting:=delRow
End Sub
