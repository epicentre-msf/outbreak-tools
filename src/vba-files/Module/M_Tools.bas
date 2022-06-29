Attribute VB_Name = "M_Tools"
Option Explicit

Sub ClicCmdAddRows(Lo As ListObject)

    'Begining of the tables
    Dim iRowHeader As Long
    Dim iColHeader  As Long

    'End of the listobject table
    Dim iRowsEnd As Long
    Dim iColsEnd As Long

    Application.EnableEvents = False
    ActiveSheet.Unprotect C_sPassword

    'Rows and columns at the begining of the table to resize
    iRowHeader = Lo.Range.Row
    iColHeader = Lo.Range.Column

    'Rows and Columns at the end of the Table to resize
    iRowsEnd = Lo.Range.Rows.Count
    iColsEnd = Lo.Range.Columns.Count

    Lo.Resize Range(Cells(iRowHeader, iColHeader), Cells(iRowsEnd + C_iNbLinesLLData, iColsEnd))

    Application.EnableEvents = True
    Call ProtectSheet

End Sub


Sub ProtectSheet()
ActiveSheet.Protect Password:=C_sPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True

End Sub


'Resize the dictionary table object
Public Sub AddRowsDict()
    Call ClicCmdAddRows(sheetDictionary.ListObjects(C_sTabDictionary))
End Sub

'Resize the choices table object
Public Sub AddRowsChoices()
    Call ClicCmdAddRows(SheetChoice.ListObjects(C_sTabChoices))
End Sub
