Attribute VB_Name = "SetupHelpers"

Option Explicit

Private Const PASSSHEETNAME As String = "__pass"
Private Const TRADSHEETNAME As String = "Translations"
Private Const ANALYSISSHEETNAME As String = "Analysis"
Private Const DICTSHEETNAME As String = "Dictionary"
Private Const UPDATEDSHEETNAME As String = "__updated"


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
    ProtectSetupSheet sheetName
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

'@sub-title Retrieve the translations listobject when available
Public Function ResolveTranslationsTable() As ListObject
    Dim sh As Worksheet

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(TRADSHEETNAME)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveTranslationsTable = sh.ListObjects(TABTRANSLATION)
    On Error GoTo 0
End Function

'@sub-title Retrieve the registry worksheet capturing updated values
Public Function ResolveRegistrySheet() As Worksheet
    On Error Resume Next
        Set ResolveRegistrySheet = ThisWorkbook.Worksheets(UPDATEDSHEETNAME)
    On Error GoTo 0
End Function

'@sub-title Clear filters on every listobject in the sheet
Public Sub ClearSheetFilters(ByVal sheetName As String)
    Dim targetSheet As Worksheet
    Dim lo As ListObject

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub

    UnProtectSetupSheet sheetName

    For Each lo In targetSheet.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            On Error Resume Next
                lo.AutoFilter.ShowAllData
            On Error GoTo 0
        End If
    Next lo

    If targetSheet.AutoFilterMode Then
        targetSheet.AutoFilterMode = False
    End If

    ProtectSetupSheet sheetName
End Sub

'@sub-title Sort setup tables based on the active worksheet
Public Sub SortSetupTables(ByVal sheetName As String)
    Dim targetSheet As Worksheet
    Dim normalizedName As String
    Dim choices As ILLChoices
    Dim analysis As IAnalysis
    Dim lo As ListObject
    Dim tabl As ICustomTable

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub

    normalizedName = LCase$(Trim$(sheetName))

    Select Case normalizedName
        Case "choices"
            Set choices = LLChoices.Create(targetSheet, START_ROW_CHOICES, START_COLUMN_CHOICES)
            UnProtectSetupSheet sheetName
                choices.Sort
            ProtectSetupSheet sheetName
        Case "analysis"
            Set analysis = Analysis.Create(targetSheet)
            UnProtectSetupSheet sheetName
                analysis.Sort
            ProtectSetupSheet sheetName
        Case "exports"
            On Error Resume Next
                Set lo = targetSheet.ListObjects(1)
            On Error GoTo 0
            If lo Is Nothing Then Exit Sub
            Set tabl = CustomTable.Create(lo)
            UnProtectSetupSheet sheetName
                tabl.Sort colName:="export number"
            ProtectSetupSheet sheetName
        Case Else
            'No action for other sheets yet
    End Select
End Sub

'@sub-title Insert a list row at the active cell position
Public Sub InsertListRowAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim position As Long

    If targetCell Is Nothing Then Exit Sub

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If Not targetCell.Parent Is targetSheet Then Exit Sub

    On Error Resume Next
        Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    UnProtectSetupSheet sheetName

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
    Else
        position = targetCell.Row - lo.HeaderRowRange.Row
        If position < 1 Or position > lo.ListRows.Count Then
            lo.ListRows.Add
        Else
            lo.ListRows.Add Position:=position
        End If
    End If

    ProtectSetupSheet sheetName
End Sub

'@sub-title Delete the list row intersecting the active cell
Public Sub DeleteListRowAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim position As Long

    If targetCell Is Nothing Then Exit Sub

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If Not targetCell.Parent Is targetSheet Then Exit Sub

    On Error Resume Next
        Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    position = targetCell.Row - lo.HeaderRowRange.Row
    If position < 1 Or position > lo.ListRows.Count Then Exit Sub

    UnProtectSetupSheet sheetName
        lo.ListRows(position).Delete
    ProtectSetupSheet sheetName
End Sub

'@sub-title Delete the list column intersecting the active cell
Public Sub DeleteListColumnAt(ByVal sheetName As String, ByVal targetCell As Range)
    Dim targetSheet As Worksheet
    Dim lo As ListObject
    Dim colIndex As Long

    If targetCell Is Nothing Then Exit Sub

    On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If Not targetCell.Parent Is targetSheet Then Exit Sub

    On Error Resume Next
        Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    colIndex = targetCell.Column - lo.Range.Column + 1
    If colIndex < 1 Or colIndex > lo.ListColumns.Count Then Exit Sub

    UnProtectSetupSheet sheetName
        lo.ListColumns(colIndex).Delete
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
